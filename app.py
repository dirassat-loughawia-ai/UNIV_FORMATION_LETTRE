import streamlit as st
from docx import Document
import base64
import re
from datetime import datetime

# 1. إعدادات الصفحة
st.set_page_config(
    page_title="منصة عروض التكوين الجامعية",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 2. التنسيق البصري للواجهة والاتجاه من اليمين للشمال (RTL)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    
    html, body, [class*="css"], [data-testid="stMarkdownContainer"] {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
    }
    
    .stApp { background-color: #f8fafc; }

    .hero-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 2rem;
        border-radius: 16px;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.3);
        margin-bottom: 2rem;
        text-align: center;
    }
    .hero-header h1 { color: #ffffff !important; font-weight: 700; margin-bottom: 0.5rem; text-align: center !important; }
    .hero-header p { color: #e0f2fe; font-size: 1rem; text-align: center !important; }

    [data-testid="stFileUploader"] {
        background: #ffffff;
        border-radius: 14px;
        padding: 1.5rem;
        border: 2px dashed #cbd5e1;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        direction: rtl !important;
        text-align: right !important;
    }

    div[data-testid="stExpander"] {
        background: #ffffff;
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        margin-bottom: 0.8rem !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02) !important;
        direction: rtl !important;
        text-align: right !important;
    }

    div[data-testid="stExpander"] summary {
        direction: rtl !important;
        text-align: right !important;
    }

    div[data-testid="stExpander"] summary p {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        text-align: right !important;
    }

    .info-card {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        padding: 1rem;
        margin-bottom: 1.2rem;
        display: flex;
        gap: 20px;
        flex-wrap: wrap;
        direction: rtl !important;
        text-align: right !important;
    }
    .info-item {
        font-weight: 600;
        color: #166534;
        font-size: 0.95rem;
    }

    .actions-container {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        margin-top: 10px;
        margin-bottom: 10px;
        direction: rtl !important;
    }

    .btn-view {
        background: #2563eb;
        color: white !important;
        padding: 8px 18px;
        border-radius: 8px;
        text-decoration: none !important;
        font-weight: 600;
        font-size: 14px;
        display: inline-block;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.2);
    }

    .btn-print {
        background: #059669;
        color: white !important;
        padding: 8px 18px;
        border-radius: 8px;
        text-decoration: none !important;
        font-weight: 600;
        font-size: 14px;
        display: inline-block;
        box-shadow: 0 2px 6px rgba(5, 150, 105, 0.2);
    }
</style>
""", unsafe_allow_html=True)

# 3. الهيدر الرئيسي
st.markdown("""
<div class="hero-header">
    <h1>🎓 دليل عروض التكوين والمناهج الجامعية</h1>
    <p>جامعة الوادي - كلية الآداب واللغات - قسم اللغة والأدب العربي</p>
</div>
""", unsafe_allow_html=True)

# 4. رفع الملف
uploaded_file = st.file_uploader("📥 قم برفع ملف Word الخاص بالتخصص لتوليد الدليل التفاعلي:", type=["docx"])

def clean_subject_name(raw_name):
    """تنظيف اسم المادة وحذف تكرار الكلمات الدليلية"""
    name = raw_name.strip()
    name = re.sub(r'^(المادة|مادّة|مادة|المقياس|مقياس|عنوان المادة)\s*[:\-]?\s*', '', name, flags=re.IGNORECASE)
    name = re.sub(r'^(مادّة|مادة)\s+', '', name, flags=re.IGNORECASE)
    return name.strip()

def clean_line_numbers(text):
    """إزالة الأرقام المبعثرة والترقيمات القديمة في أول أو نهاية السطر"""
    t = text.strip()
    t = re.sub(r'^[\(\[\{]?\d+[\)\]\}]?\s*[\.\-\:]?\s*', '', t)
    t = re.sub(r'\s+[\(\[\{]?\d+[\)\]\}]?\s*$', '', t)
    return t.strip()

def parse_docx_curriculum(file):
    doc = Document(file)
    semesters = {}
    
    raw_blocks = []
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            raw_blocks.append(txt)
            
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    txt = p.text.strip()
                    if txt and txt not in raw_blocks:
                        raw_blocks.append(txt)

    current_sem = None
    current_sub = None
    expecting_subject_after_sem = False
    
    def get_template():
        return {
            "unit": "وحدة التعليم الأساسية",
            "coefficient": "02",
            "credits": "05",
            "topics": [],
            "references": []
        }

    ignore_keywords = ["محتوى", "برنامج", "أهداف", "محور", "فصل", "مقدمة", "خاتمة", "تمهيد", "عنصر", "المحور", "الفصل"]
    ref_keywords = ["المراجع", "مراجع", "المصادر", "مصادر", "قائمة المراجع", "قائمة المصادر", "المراجع والمصادر", "المصادر والمراجع"]

    for text in raw_blocks:
        # 1. رصد السداسي
        if "السداسي" in text and len(text) < 40:
            sem_m = re.search(r'(السداسي\s+[\u0600-\u06FF]+)', text)
            current_sem = sem_m.group(1) if sem_m else text
            if current_sem not in semesters:
                semesters[current_sem] = {}
            current_sub = None
            expecting_subject_after_sem = True
            continue

        # 2. رصد المادة الجديدة
        is_subject_line = any(text.startswith(k) or re.match(r'^(المادة|مادّة|مادة|المقياس|مقياس)\s*[:\-]', text) for k in ["المادة:", "مادّة:", "مادة:", "المقياس:", "مقياس:"])
        
        if current_sem and (is_subject_line or "عنوان المادة" in text or expecting_subject_after_sem):
            if not any(ik in text for ik in ignore_keywords) and len(text) < 120:
                sub_m = re.search(r'(مادّة|مادة|مقياس|عنوان المادة)\s*[:\-]?\s*([\u0600-\u06FF\s]+?)(?=\s+(المعامل|الرصيد|وحدة|$))', text)
                extracted_name = sub_m.group(2) if sub_m else text

                clean_name = clean_subject_name(extracted_name)
                
                if clean_name and len(clean_name) > 2 and "السداسي" not in clean_name:
                    current_sub = clean_name
                    if current_sub not in semesters[current_sem]:
                        semesters[current_sem][current_sub] = get_template()
                    
                    expecting_subject_after_sem = False

                    u_m = re.search(r'(وحدة\s+التّعليم\s+[\u0600-\u06FF]+|وحدة\s+[\u0600-\u06FF]+)', text)
                    c_m = re.search(r'المعامل\s*[:\-]?\s*(\d+)', text)
                    r_m = re.search(r'الرصيد\s*[:\-]?\s*(\d+)', text)

                    if u_m: semesters[current_sem][current_sub]["unit"] = u_m.group(1).strip()
                    if c_m: semesters[current_sem][current_sub]["coefficient"] = c_m.group(1).strip()
                    if r_m: semesters[current_sem][current_sub]["credits"] = r_m.group(1).strip()
                    continue

        # 3. تصنيف الأسطر إلى محتوى أو مراجع بشكل منفصل
        if current_sem and current_sub and current_sub in semesters[current_sem]:
            cleaned_txt = clean_line_numbers(text)
            if cleaned_txt:
                is_ref_header = any(cleaned_txt == rk or cleaned_txt.startswith(rk + ":") or cleaned_txt.startswith(rk + " ") for rk in ref_keywords)
                
                if is_ref_header:
                    semesters[current_sem][current_sub]["is_in_refs"] = True
                    continue

                if semesters[current_sem][current_sub].get("is_in_refs", False):
                    semesters[current_sem][current_sub]["references"].append(cleaned_txt)
                else:
                    semesters[current_sem][current_sub]["topics"].append(cleaned_txt)

    return semesters

# دالة توليد وثيقة PDF رسمية بتنسيق احترافي
def generate_pdf_html(sem_title, sub_title, data):
    today_date = datetime.now().strftime("%Y/%m/%d")
    
    if data["topics"]:
        topics_html = "".join([f"<li>{item}</li>" for item in data["topics"]])
    else:
        topics_html = "<li>محتوى المادة مسجل في الوثيقة المصدر.</li>"

    if data["references"]:
        refs_list_html = "".join([f"<li>{item}</li>" for item in data["references"]])
        refs_section_html = f"""
        <div class="section-block">
            <h3 class="section-header">📚 قائمة المصادر والمراجع:</h3>
            <ol class="custom-list refs-list">
                {refs_list_html}
            </ol>
        </div>
        """
    else:
        refs_section_html = ""

    html_code = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>{sub_title}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
        <style>
            @page {{
                size: A4;
                margin-top: 1.5cm;
                margin-bottom: 1.5cm;
                margin-left: 1.5cm;
                margin-right: 1.5cm;
                @bottom-center {{
                    content: counter(page) " - " counter(pages);
                    font-family: 'Cairo', sans-serif;
                    font-size: 11px;
                    font-weight: 700;
                    color: #64748b;
                }}
            }}

            html, body {{
                margin: 0;
                padding: 0;
                background-color: #ffffff;
            }}

            body {{ 
                font-family: 'Cairo', sans-serif; 
                color: #0f172a; 
                direction: rtl; 
                text-align: right; 
                line-height: 1.7;
                font-size: 14px;
            }}

            /* حاوية الورقة مع حواشي أمان صريحة */
            .page-container {{
                box-sizing: border-box;
                padding: 0.5cm 0; /* حاشية أمان مضافة للداخل */
                width: 100%;
            }}

            /* الترويسة الرسمية */
            .official-header {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 2px solid #1e3a8a;
                padding-bottom: 12px;
            }}
            .center-title-1 {{
                font-size: 14px;
                font-weight: 700;
                margin: 0;
                color: #0f172a;
            }}
            .center-title-2 {{
                font-size: 16px;
                font-weight: 800;
                margin: 4px 0 10px 0;
                color: #1e3a8a;
            }}
            .header-flex {{
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                font-size: 13px;
                font-weight: 700;
                color: #334155;
            }}
            .header-right {{ text-align: right; }}
            .header-left {{ text-align: left; }}

            /* عنوان المادة والسداسي */
            .doc-title-box {{
                background-color: #f8fafc;
                border: 1px solid #cbd5e1;
                border-right: 6px solid #1e3a8a;
                padding: 12px 18px;
                border-radius: 6px;
                margin-bottom: 18px;
            }}
            .doc-title-box h2 {{ font-size: 13px; color: #475569; margin: 0 0 4px 0; }}
            .doc-title-box h1 {{ font-size: 19px; color: #1e3a8a; margin: 0; font-weight: 800; }}

            /* بطاقة المعلمات */
            .meta-grid {{
                display: flex;
                gap: 20px;
                background-color: #f1f5f9;
                padding: 10px 16px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: 700;
                margin-bottom: 22px;
            }}
            .meta-item {{ color: #1e293b; }}

            /* الأقسام وتدفق المحتوى */
            .section-block {{
                break-inside: auto !important;
                page-break-inside: auto !important;
                margin-bottom: 18px;
            }}

            .section-header {{
                color: #1e3a8a;
                font-size: 16px;
                font-weight: 800;
                border-bottom: 2px solid #e2e8f0;
                padding-bottom: 6px;
                margin-top: 18px;
                margin-bottom: 14px;
            }}

            /* الترقيم والمقروئية */
            ol.custom-list {{
                list-style: none;
                counter-reset: topic-counter;
                padding-right: 0;
                margin-top: 8px;
                margin-bottom: 8px;
                break-inside: auto !important;
                page-break-inside: auto !important;
            }}
            
            ol.custom-list li {{
                counter-increment: topic-counter;
                position: relative;
                padding-right: 38px;
                margin-bottom: 12px;
                font-size: 14px;
                line-height: 1.7;
                text-align: justify;
                color: #1e293b;
                break-inside: avoid-page !important;
                page-break-inside: avoid !important;
            }}
            
            ol.custom-list li::before {{
                content: counter(topic-counter);
                position: absolute;
                right: 0;
                top: 2px;
                width: 24px;
                height: 24px;
                background-color: #2563eb;
                color: #ffffff;
                font-weight: 700;
                font-size: 12px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
            }}

            /* ترقيم المراجع */
            ol.refs-list {{
                counter-reset: ref-counter !important;
            }}
            ol.refs-list li {{
                counter-increment: ref-counter !important;
            }}
            ol.refs-list li::before {{
                content: counter(ref-counter) !important;
                background-color: #059669 !important;
            }}

            .footer-date {{
                margin-top: 30px;
                padding-top: 10px;
                border-top: 1px dashed #cbd5e1;
                text-align: left;
                font-size: 12px;
                font-weight: 700;
                color: #64748b;
                break-inside: avoid;
            }}

            @media screen {{
                body {{
                    background-color: #f1f5f9;
                    padding: 20px;
                }}
                .page-container {{
                    background: #ffffff;
                    max-width: 210mm;
                    margin: 0 auto;
                    padding: 1.5cm;
                    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
                    border-radius: 4px;
                }}
            }}

            @media print {{
                .no-print {{ display: none !important; }}
                body {{
                    background-color: #ffffff !important;
                    padding: 0 !important;
                    margin: 0 !important;
                }}
                .page-container {{
                    padding: 0 !important;
                    margin: 0 !important;
                    box-shadow: none !important;
                    width: 100% !important;
                }}
            }}

            .print-btn {{
                background: #059669;
                color: white;
                border: none;
                padding: 10px 24px;
                border-radius: 6px;
                font-weight: 700;
                cursor: pointer;
                margin-bottom: 20px;
                font-size: 15px;
                box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            }}
        </style>
    </head>
    <body>
        <button class="print-btn no-print" onclick="window.print()">🖨️ طباعة أو حفظ الوثيقة (PDF)</button>
        
        <div class="page-container">
            <!-- الترويسة الرسمية -->
            <div class="official-header">
                <div class="center-title-1">الجمهورية الجزائرية الديمقراطية الشعبية</div>
                <div class="center-title-1">وزارة التعليم العالي والبحث العلمي</div>
                <div class="center-title-2">جامعة الوادي</div>
                
                <div class="header-flex">
                    <div class="header-right">
                        <div>كلية الآداب واللغات</div>
                        <div style="margin-top: 4px; color: #1e3a8a;">قسم اللغة والأدب العربي</div>
                    </div>
                    <div class="header-left">
                        <div>إدارة الشعبتين اللغوية والأدبية</div>
                        <div style="margin-top: 4px; color: #1e3a8a;">السنة الجامعية: 2026-2027</div>
                    </div>
                </div>
            </div>

            <!-- عنوان المادة والبيانات -->
            <div class="doc-title-box">
                <h2>{sem_title}</h2>
                <h1>المادة: {sub_title}</h1>
            </div>

            <div class="meta-grid">
                <div class="meta-item">📌 {data['unit']}</div>
                <div class="meta-item">⚖️ المعامل: {data['coefficient']}</div>
                <div class="meta-item">🎯 الرصيد: {data['credits']}</div>
            </div>

            <!-- مفردات المادة -->
            <div class="section-block">
                <h3 class="section-header">📘 مفردات برنامج المادة:</h3>
                <ol class="custom-list">
                    {topics_html}
                </ol>
            </div>

            <!-- قائمة المراجع -->
            {refs_section_html}

            <!-- التاريخ -->
            <div class="footer-date">
                📅 حرر بتاريخ: {today_date} م
            </div>
        </div>
    </body>
    </html>
    """
    b64 = base64.b64encode(html_code.encode('utf-8')).decode('utf-8')
    return f"data:text/html;charset=utf-8;base64,{b64}"

# 5. عرض البيانات المنسقة
if uploaded_file is not None:
    data = parse_docx_curriculum(uploaded_file)
    st.success("✨ تم تطبيق هوامش 1.5 سم وتوفير مسافات مريحة للقراءة والطباعة!")

    for sem_title, subjects in data.items():
        if not subjects:
            continue
        with st.expander(f"📌 {sem_title}", expanded=False):
            for sub_title, sub_data in subjects.items():
                with st.expander(f"📖 المادة: {sub_title}", expanded=False):
                    
                    st.markdown(f"""
                    <div class="info-card">
                        <div class="info-item">🏛️ <b>الوحدة:</b> {sub_data['unit']}</div>
                        <div class="info-item">⚖️ <b>المعامل:</b> {sub_data['coefficient']}</div>
                        <div class="info-item">🎯 <b>الرصيد:</b> {sub_data['credits']}</div>
                    </div>
                    """, unsafe_allow_html=True)

                    pdf_link = generate_pdf_html(sem_title, sub_title, sub_data)
                    
                    st.markdown(f"""
                    <div class="actions-container">
                        <a href="{pdf_link}" target="_blank" class="btn-view">👁️ معاينة برنامج المادة</a>
                        <a href="{pdf_link}" download="{sub_title}.html" class="btn-print">📥 تحميل الملف / طباعة PDF</a>
                    </div>
                    """, unsafe_allow_html=True)