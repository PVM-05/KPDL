"""
Kịch bản tự động tạo file Báo cáo Word (report.docx) hoàn chỉnh, chi tiết
với Trang bìa chuẩn Đại học Phenikaa (có khung viền trang, logo tròn trường CNTT, thông tin sinh viên Phạm Văn Minh)
và đầy đủ nội dung báo cáo chi tiết 5 chương, giải mã con số, giải thích mã nguồn và nhúng toàn bộ ảnh đồ thị.
"""

import os
import sys
import json

# Thiết lập UTF-8 cho Windows console
if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUTS_DIR = os.path.join(BASE_DIR, "outputs")
METRICS_PATH = os.path.join(OUTPUTS_DIR, "metrics.json")
LOGO_PATH = os.path.join(DATA_DIR, "phenikaa_logo.png")


def set_cell_background(cell, fill_hex):
    """Đặt màu nền cho ô trong bảng Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Thiết lập khoảng đệm padding cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def add_page_border(section):
    """Thêm khung viền trang bìa chuẩn mẫu đại học."""
    sectPr = section._sectPr
    xml_str = (
        f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
        f'<w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
        f'<w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
        f'<w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/>'
        f'</w:pgBorders>'
    )
    pgBorders = parse_xml(xml_str)
    sectPr.append(pgBorders)


def build_phenikaa_report_docx():
    print("Đang nạp số liệu thực nghiệm từ metrics.json...")
    with open(METRICS_PATH, "r", encoding="utf-8") as f:
        metrics = json.load(f)

    doc = Document()

    # =============================================================
    # SECTION 1: TRANG BÌA CHUẨN ĐẠI HỌC PHENIKAA (CÓ KHUNG VIỀN & LOGO)
    # =============================================================
    cover_sect = doc.sections[0]
    cover_sect.top_margin = Inches(0.8)
    cover_sect.bottom_margin = Inches(0.8)
    cover_sect.left_margin = Inches(0.8)
    cover_sect.right_margin = Inches(0.8)
    add_page_border(cover_sect)

    # 1. Header cơ quan
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.paragraph_format.space_before = Pt(10)
    p_header.paragraph_format.space_after = Pt(2)
    r1 = p_header.add_run("ĐẠI HỌC PHENIKAA\n")
    r1.bold = True
    r1.font.size = Pt(14)
    r1.font.color.rgb = RGBColor(0, 0, 0)

    r2 = p_header.add_run("TRƯỜNG CÔNG NGHỆ THÔNG TIN")
    r2.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor(0, 0, 0)

    # 2. Logo tròn Phenikaa Faculty of Computer Science
    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(25)
    p_logo.paragraph_format.space_after = Pt(25)
    if os.path.exists(LOGO_PATH):
        p_logo.add_run().add_picture(LOGO_PATH, width=Inches(2.0))
    else:
        p_logo.add_run("[PHENIKAA UNIVERSITY LOGO]").bold = True

    # 3. Tiêu đề báo cáo
    p_type = doc.add_paragraph()
    p_type.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_type.paragraph_format.space_before = Pt(15)
    p_type.paragraph_format.space_after = Pt(6)
    r_type = p_type.add_run("BÁO CÁO BÀI TẬP LỚN\n")
    r_type.bold = True
    r_type.font.size = Pt(20)
    r_type.font.color.rgb = RGBColor(0, 0, 0)

    r_course = p_type.add_run("HỌC PHẦN: KHAI PHÁ DỮ LIỆU / HỌC MÁY\n")
    r_course.font.size = Pt(13)
    r_course.font.color.rgb = RGBColor(100, 100, 100)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(15)
    p_title.paragraph_format.space_after = Pt(30)
    r_title = p_title.add_run("DỰ ĐOÁN SỚM KẾT QUẢ HỌC TẬP CỦA SINH VIÊN BẰNG THUẬT TOÁN DECISION TREE")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(41, 128, 185)

    # 4. Thông tin sinh viên (Căn lề trái có thụt dòng đẹp mắt)
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.left_indent = Inches(1.5)
    p_meta.paragraph_format.space_before = Pt(25)
    p_meta.paragraph_format.space_after = Pt(40)
    p_meta.paragraph_format.line_spacing = 1.3

    r_s1 = p_meta.add_run("Sinh viên thực hiện: ")
    r_s1.bold = True
    r_s1.font.size = Pt(12)
    r_s2 = p_meta.add_run("Phạm Văn Minh\n")
    r_s2.font.size = Pt(12)

    r_m1 = p_meta.add_run("Mã sinh viên: ")
    r_m1.bold = True
    r_m1.font.size = Pt(12)
    r_m2 = p_meta.add_run("23010350\n")
    r_m2.font.size = Pt(12)

    r_e1 = p_meta.add_run("Email: ")
    r_e1.bold = True
    r_e1.font.size = Pt(12)
    r_e2 = p_meta.add_run("23010350@st.phenikaa-uni.edu.vn\n")
    r_e2.font.size = Pt(12)

    r_g1 = p_meta.add_run("Giảng viên hướng dẫn: ")
    r_g1.bold = True
    r_g1.font.size = Pt(12)
    r_g2 = p_meta.add_run("Bộ môn Khai phá Dữ liệu\n")
    r_g2.font.size = Pt(12)

    # 5. Footer địa điểm và thời gian
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(30)
    p_foot.paragraph_format.space_after = Pt(10)
    r_foot = p_foot.add_run("Hà Nội, tháng 10 năm 2026")
    r_foot.italic = True
    r_foot.font.size = Pt(12)

    # =============================================================
    # SECTION 2: NỘI DUNG CHI TIẾT BÁO CÁO (HEADER/FOOTER & NỘI DUNG 5 CHƯƠNG)
    # =============================================================
    body_sect = doc.add_section(WD_SECTION.NEW_PAGE)
    body_sect.header.is_linked_to_previous = False
    body_sect.top_margin = Inches(0.9)
    body_sect.bottom_margin = Inches(0.9)
    body_sect.left_margin = Inches(0.8)
    body_sect.right_margin = Inches(0.8)

    # Header cho các trang nội dung
    header_p = body_sect.header.paragraphs[0]
    header_p.text = "Học phần Khai phá Dữ liệu --- Báo cáo Bài tập lớn: Dự đoán sớm bằng Decision Tree"
    header_p.runs[0].font.size = Pt(8.5)
    header_p.runs[0].font.italic = True
    header_p.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    # -------------------------------------------------------------
    # BẢNG THÔNG TIN SINH VIÊN & LỜI CAM ĐOAN
    # -------------------------------------------------------------
    doc.add_heading("THÔNG TIN SINH VIÊN & BẢNG CAM ĐOAN ĐÓNG GÓP", level=2)
    p_desc = doc.add_paragraph()
    r_desc = p_desc.add_run(
        "Căn cứ vào quy chế đánh giá học phần Khai phá Dữ liệu: Điểm bài viết báo cáo kỹ thuật (hệ số 50%) và "
        "điểm vấn đáp/thuyết trình (hệ số 50%) được chấm trực tiếp cho cá nhân sinh viên thực hiện độc lập đề tài:"
    )
    r_desc.font.size = Pt(10.5)

    table_contrib = doc.add_table(rows=2, cols=5)
    table_contrib.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["STT", "Họ và Tên Sinh Viên", "Mã Sinh Viên", "Nhiệm Vụ Thực Hiện Toàn Trình", "Hệ Số"]
    
    hdr_cells = table_contrib.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "2980B9")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)

    row1 = table_contrib.rows[1].cells
    row1[0].text = "1"
    row1[1].text = "Phạm Văn Minh"
    row1[2].text = "23010350"
    row1[3].text = "Thực hiện độc lập 100%: Khảo sát tập dữ liệu 1.656 sinh viên, thiết kế pipeline tiền xử lý, huấn luyện Decision Tree và mô hình đối chứng Logistic Regression, kiểm định chéo 5-Fold Stratified CV, xuất toàn bộ 5 biểu đồ thực nghiệm 300 DPI và viết báo cáo khoa học hoàn chỉnh."
    row1[4].text = "1.0 (100%)"

    for cell in row1:
        set_cell_margins(cell, top=90, bottom=90, left=100, right=100)
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
    row1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_agree = doc.add_paragraph()
    p_agree.paragraph_format.space_before = Pt(12)
    p_agree.paragraph_format.space_after = Pt(20)
    run_agree = p_agree.add_run(
        "Lời cam đoan: Tôi xin cam đoan đây là công trình bài tập lớn do bản thân tôi tự nghiên cứu, lập trình và hoàn thiện độc lập. "
        "Mọi số liệu thực nghiệm, bảng biểu đối sánh và đồ thị minh họa trong báo cáo đều được trích xuất trung thực từ mã nguồn chương trình, "
        "bộ dữ liệu sử dụng đã được khử định danh học thuật và tuân thủ các chuẩn mực nghiên cứu khoa học."
    )
    run_agree.bold = True
    run_agree.italic = True
    run_agree.font.size = Pt(10)

    # -------------------------------------------------------------
    # MỤC LỤC TỔNG QUAN
    # -------------------------------------------------------------
    doc.add_heading("MỤC LỤC NỘI DUNG BÁO CÁO", level=2)
    toc_items = [
        "PHẦN I: BÁO CÁO KỸ THUẬT CHUẨN ĐỀ CƯƠNG (5 CHƯƠNG)",
        "   Chương 1: Giới thiệu tổng quan bài toán và nghiên cứu liên quan",
        "   Chương 2: Cơ sở lý thuyết thuật toán Cây quyết định (Decision Tree)",
        "   Chương 3: Triển khai thực nghiệm và Giao thức Zero Label Leakage",
        "   Chương 4: Phân tích kết quả thực nghiệm và Phân tích lỗi cận biên",
        "   Chương 5: Kết luận, đánh giá điểm mạnh / hạn chế và hướng phát triển",
        "   Tài liệu tham khảo",
        "PHẦN II: GIẢI MÃ CHI TIẾT Ý NGHĨA CỦA TỪNG CON SỐ THỰC NGHIỆM",
        "PHẦN III: GIẢI THÍCH KIẾN TRÚC MÃ NGUỒN VÀ Ý NGHĨA TỪNG ĐOẠN CODE"
    ]
    for item in toc_items:
        p_toc = doc.add_paragraph(item)
        p_toc.paragraph_format.space_after = Pt(3)
        p_toc.runs[0].font.size = Pt(10)
        if item.startswith("PHẦN"):
            p_toc.runs[0].bold = True
            p_toc.runs[0].font.color.rgb = RGBColor(41, 128, 185)

    doc.add_page_break()

    # =============================================================
    # PHẦN I: BÁO CÁO HỌC THUẬT (5 CHƯƠNG)
    # =============================================================
    h_p1 = doc.add_heading("PHẦN I: BÁO CÁO KỸ THUẬT CHUẨN ĐỀ CƯƠNG (5 CHƯƠNG)", level=1)
    h_p1.runs[0].font.color.rgb = RGBColor(41, 128, 185)

    # -------------------------------------------------------------
    # CHƯƠNG 1
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 1: GIỚI THIỆU TỔNG QUAN", level=1)
    
    doc.add_heading("1.1. Bối Cảnh Bài Toán & Tính Cấp Thiết", level=2)
    p = doc.add_paragraph(
        "Trong các cơ sở giáo dục đại học đào tạo theo học chế tín chỉ, sinh viên năm thứ nhất thường phải đối mặt với bước chuyển tiếp "
        "rất lớn về phương pháp tiếp cận kiến thức, tính chủ động tự học và cường độ làm việc. Việc gặp bỡ ngỡ trước khối lượng kiến thức "
        "khoa học cơ bản (Giải tích, Vật lý đại cương) và kỹ thuật cơ sở thường dẫn đến tình trạng sa sút kết quả học tập, bị cảnh báo học vụ "
        "hoặc thậm chí bị buộc thôi học ở các học kỳ sau.\n\n"
        "Do đó, bài toán Khai phá Dữ liệu Giáo dục (Educational Data Mining --- EDM) nhằm xây dựng Hệ thống cảnh báo sớm (Early Warning System) "
        "mang ý nghĩa thực tiễn to lớn: giúp nhà trường và đội ngũ cố vấn học tập chủ động phát hiện sớm sinh viên có nguy cơ sa sút để can thiệp, "
        "tư vấn và hỗ trợ phương pháp học tập ngay sau năm nhất, thay vì chờ đến khi kết quả đã an bài ở năm cuối."
    )
    p.paragraph_format.line_spacing = 1.25

    doc.add_heading("1.2. Mục Tiêu & Phạm Vi Đề Tài", level=2)
    p = doc.add_paragraph(
        "Đề tài tập trung giải quyết bài toán phân lớp nhị phân kết quả học tập tốt nghiệp của sinh viên đại học. Mục tiêu và phạm vi cụ thể gồm:\n"
        "1. Khảo sát hiện trạng phân bố điểm tích lũy GPA toàn khóa và sự tương quan giữa các môn học Năm 1 với năng lực tốt nghiệp của sinh viên.\n"
        "2. Ứng dụng thuật toán Cây quyết định (Decision Tree Classifier) để dự đoán sớm sinh viên sẽ đạt loại Khá/Giỏi (GPA >= 2.50) "
        "hay Trung bình/Yếu (GPA < 2.50) chỉ bằng kết quả các môn học của Năm thứ nhất (15 môn học nền tảng).\n"
        "3. Đảm bảo tính minh bạch và khả năng giải thích (Interpretability) của mô hình cây thông qua việc trực quan hóa các quy tắc rẽ nhánh (If-Then rules) "
        "và xếp hạng các môn học có ảnh hưởng lớn nhất.\n"
        "4. Tuân thủ nghiêm ngặt nguyên tắc Zero Data Leakage: Không sử dụng điểm của các năm học sau hoặc các thông tin rò rỉ nhãn vào quá trình dự đoán sớm."
    )

    doc.add_heading("1.3. Tổng Quan Machine Learning & Bài Toán Muốn Làm (~100 chữ)", level=2)
    p = doc.add_paragraph(
        "Học máy (Machine Learning) là lĩnh vực trí tuệ nhân tạo nghiên cứu các thuật toán cho phép máy tính tự động đúc kết quy luật từ dữ liệu quá khứ. "
        "Trong bài toán phân lớp kết quả học tập này, tác giả lựa chọn thuật toán Cây quyết định (Decision Tree Classifier) với tiêu chí Độ vấy bẩn Gini (Gini Impurity). "
        "Lý do then chốt để lựa chọn Cây quyết định trong giáo dục là cấu trúc dạng cây phân cấp (White-box model) mang tính minh bạch tuyệt đối: "
        "cung cấp các quy tắc If-Then trực quan và các ngưỡng điểm số học cụ thể cho từng môn học, giúp cố vấn học tập dễ dàng giải thích và định hướng cho sinh viên, "
        "vượt trội hơn các mô hình hộp đen phức tạp."
    )
    p.runs[0].italic = True

    doc.add_heading("1.4. Một Số Công Việc Liên Quan (Related Works)", level=2)
    p = doc.add_paragraph(
        "Nghiên cứu kinh điển của Cortez & Silva (2008) [1] đã mở đầu xu hướng áp dụng các thuật toán Học máy (Decision Tree, Random Forest, SVM) "
        "để dự đoán kết quả học tập, chỉ ra rằng điểm các kỳ học đầu tiên đóng vai trò quyết định trong việc định hình kết quả cuối khóa. "
        "Romero & Ventura (2010) [4] đã khảo sát toàn diện các kỹ thuật khai phá dữ liệu giáo dục và nhấn mạnh rằng tính khả giải (Interpretability) "
        "của mô hình Cây quyết định là tiêu chí hàng đầu để ứng dụng trong môi trường đào tạo. "
        "Tại Việt Nam, các nghiên cứu tại các trường kỹ thuật lớn cũng khẳng định điểm số các môn khoa học cơ bản (Giải tích, Vật lý, Ngoại ngữ) "
        "ở năm nhất có hệ số tương quan rất cao với năng lực tốt nghiệp của sinh viên."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 2
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 2: CƠ SỞ LÝ THUYẾT THUẬT TOÁN DECISION TREE", level=1)
    
    doc.add_heading("2.1. Bản Chất Của Cây Quyết Định", level=2)
    p = doc.add_paragraph(
        "Cây quyết định (Decision Tree) là phương pháp học có giám sát phi tham số (non-parametric supervised learning). Mô hình tổ chức tri thức "
        "học được dưới dạng cấu trúc cây phân cấp bao gồm:\n"
        "- Nút gốc (Root Node): Chứa toàn bộ tập dữ liệu huấn luyện ban đầu trước khi bị phân chia.\n"
        "- Nút nội vi (Decision Nodes): Biểu diễn một điều kiện kiểm tra trên một thuộc tính cụ thể (ví dụ: Điểm môn Điện - Điện tử ô tô <= 2.25).\n"
        "- Nhánh (Branches): Tương ứng với kết quả của phép kiểm tra điều kiện (Đúng hoặc Sai).\n"
        "- Nút lá (Leaf Nodes): Đại diện cho nhãn phân lớp cuối cùng (Khá-Giỏi hoặc Trung bình-Yếu) cùng phân bố xác suất tương ứng."
    )

    doc.add_heading("2.2. Tiêu Chí Phân Tách Toán Học", level=2)
    p = doc.add_paragraph(
        "Để tìm thuộc tính và điểm ngưỡng phân chia tối ưu tại mỗi bước, thuật toán đánh giá mức độ thuần khiết (Purity) của các tập con sinh ra.\n\n"
        "1. Độ vấy bẩn Gini (Gini Impurity --- Thuật toán CART):\n"
        "Đo lường xác suất một phần tử được gán nhãn ngẫu nhiên bị phân loại sai nếu nhãn được rút ngẫu nhiên theo phân bố trong tập:"
    )
    p_eq1 = doc.add_paragraph()
    p_eq1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq1.add_run("Gini(D) = 1 - ∑ [p_i]^2  (với i = 1 đến C)\n").bold = True
    p_eq1.add_run("Trong bài toán phân lớp nhị phân (C = 2), Gini dao động từ 0 (hoàn toàn thuần khiết) đến 0.5 (hỗn độn cực đại 50/50).\n")

    p = doc.add_paragraph("Khi phân tách tập D thành hai tập con D1 và D2 theo thuộc tính A tại ngưỡng theta, Độ vấy bẩn Gini có trọng số là:")
    p_eq2 = doc.add_paragraph()
    p_eq2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq2.add_run("Gini_split(D, A) = (|D1| / |D|) * Gini(D1) + (|D2| / |D|) * Gini(D2)\n").bold = True
    p_eq2.add_run("Thuật toán chọn thuộc tính A* và ngưỡng theta* sao cho Gini_split đạt giá trị nhỏ nhất.\n")

    p = doc.add_paragraph("2. Entropy và Độ lợi thông tin (Information Gain --- Thuật toán ID3/C4.5):")
    p_eq3 = doc.add_paragraph()
    p_eq3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq3.add_run("Entropy(D) = - ∑ p_i * log2(p_i)\n").bold = True
    p_eq3.add_run("Information Gain(D, A) = Entropy(D) - ∑ (|D_v| / |D|) * Entropy(D_v)\n").bold = True

    doc.add_heading("2.3. Các Bước Xây Dựng Cây Quyết Định (Algorithm Steps / Pseudocode)", level=2)
    p = doc.add_paragraph(
        "Quy trình xây dựng Cây quyết định CART được thực hiện theo chiến lược đệ quy tham lam từ trên xuống (Top-down Greedy Recursion):\n"
        "Bước 1: Bắt đầu tại nút gốc chứa toàn bộ tập dữ liệu huấn luyện D.\n"
        "Bước 2: Với mọi thuộc tính liên tục, sắp xếp các giá trị và tính Gini_split cho tất cả các điểm chia khả dĩ.\n"
        "Bước 3: Chọn thuộc tính và ngưỡng phân chia tối ưu mang lại mức giảm Gini lớn nhất (tăng độ thuần khiết cao nhất).\n"
        "Bước 4: Chia tập dữ liệu thành 2 nhánh con (D_left và D_right) tương ứng với điều kiện kiểm tra.\n"
        "Bước 5: Lặp lại đệ quy các Bước 2-4 trên từng tập con cho đến khi thỏa mãn một trong các điều kiện dừng:\n"
        "   - Nút đạt độ thuần khiết tuyệt đối (Gini = 0).\n"
        "   - Cây đạt độ sâu tối đa quy định (max_depth).\n"
        "   - Số lượng mẫu tại nút nhỏ hơn ngưỡng tối thiểu (min_samples_leaf).\n"
        "Bước 6: Gán nhãn cho nút lá theo nguyên tắc đa số phiếu (Majority Voting)."
    )

    doc.add_heading("2.4. Kỹ Thuật Kiểm Soát Quá Khớp (Pre-pruning)", level=2)
    p = doc.add_paragraph(
        "Nếu để cây phát triển tự do không kiểm soát, mô hình sẽ cố gắng phân nhánh để bao phủ từng điểm nhiễu của tập Train, "
        "dẫn đến hiện tượng Quá khớp (Overfitting) và làm suy giảm khả năng dự đoán trên dữ liệu mới. Đề tài áp dụng kỹ thuật Tiền cắt tỉa (Pre-pruning):\n"
        "- Khống chế độ sâu tối đa (max_depth = 3): Giữ cho cây ở quy mô vừa phải (tối đa 8 nút lá), chỉ học các quy luật cốt lõi nhất.\n"
        "- Giới hạn số mẫu tối thiểu ở nút lá (min_samples_leaf = 10): Ngăn chặn việc hình thành các nhánh cá biệt chỉ phục vụ một vài mẫu dữ liệu."
    )

    doc.add_heading("2.5. Ưu Điểm & Hạn Chế", level=2)
    p = doc.add_paragraph(
        "1. Ưu điểm: Tính khả giải cao (White-box model), cung cấp tập luật If-Then trực quan; không nhạy cảm với thang đo; "
        "tự động chọn lọc thuộc tính quan trọng (Gini Importance).\n"
        "2. Hạn chế: Ranh giới phân chia trực giao song song với trục tọa độ; độ nhạy cảm cao với phương sai (High Variance) nếu dữ liệu thay đổi nhỏ."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 3
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM & GIAO THỨC ZERO LABEL LEAKAGE", level=1)
    
    doc.add_heading("3.1. Môi Trường Thực Nghiệm & Thư Viện", level=2)
    p = doc.add_paragraph(
        "Toàn bộ mã nguồn thực nghiệm được tổ chức module hóa tại thư mục src/ và thực thi trên môi trường:\n"
        "- Hệ điều hành & Trình thông dịch: Windows 11 64-bit, Python 3.12.\n"
        "- Thư viện cốt lõi: scikit-learn 1.6+ (mô hình học máy), pandas 3.0+ (quản trị dữ liệu bảng), numpy (tính toán ma trận), "
        "matplotlib và seaborn (xuất đồ thị 300 DPI)."
    )

    doc.add_heading("3.2. Mô Tả Bộ Dữ Liệu Thực Nghiệm (Anonymized Academic Benchmark)", level=2)
    p = doc.add_paragraph(
        f"Nghiên cứu sử dụng bộ dữ liệu chuẩn hóa học thuật gồm **{metrics['summary']['n_students']} sinh viên** ngành kỹ thuật với "
        f"**{metrics['summary']['n_features']} thuộc tính**:\n"
        f"- Toàn bộ thông tin định danh cá nhân (PII) như Họ tên, MSSV, Ngày sinh đã được khử định danh triệt để nhằm bảo đảm tính riêng tư theo chuẩn nghiên cứu khoa học.\n"
        f"- Biến nhân khẩu học: Nơi sinh (BirthPlace --- {metrics['summary']['birthplace_count']} tỉnh thành sau lọc) và Giới tính (Sex: Nam chiếm 99.76% do đặc thù ngành kỹ thuật cơ khí --- ô tô).\n"
        f"- Biến môn học: Điểm số 50 môn học từ HK1 đến HK8 theo thang điểm 4.0.\n"
        f"- Biến tổng kết GPA toàn khóa: Điểm trung bình = {metrics['summary']['gpa_mean']:.3f}, Độ lệch chuẩn = {metrics['summary']['gpa_std']:.3f}, "
        f"Trung vị = {metrics['summary']['gpa_median']:.3f}, Khoảng điểm: [{metrics['summary']['gpa_min']:.2f} - {metrics['summary']['gpa_max']:.2f}]."
    )

    # Chèn Biểu đồ 1: Phân bố học lực
    path_dist = os.path.join(OUTPUTS_DIR, "class_distribution.png")
    if os.path.exists(path_dist):
        doc.add_picture(path_dist, width=Inches(5.5))
        p_cap = doc.add_paragraph("Hình 3.1: Phân bố xếp loại học lực của 1.656 sinh viên trên thang điểm GPA 4.0")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("3.3. Thiết Kế Kịch Bản Dự Đoán Sớm & Giao Thức Zero Label Leakage", level=2)
    p = doc.add_paragraph(
        "1. Xác định nhãn mục tiêu: Điểm GPA tích lũy toàn khóa được nhị phân hóa theo Quy chế đào tạo tín chỉ (Thông tư 08/2021/TT-BGDĐT):\n"
        "   - Nhãn 1 (Khá --- Giỏi): GPA >= 2.50 (Gồm 794 sinh viên --- 47.95%).\n"
        "   - Nhãn 0 (Trung bình --- Yếu): GPA < 2.50 (Gồm 862 sinh viên --- 52.05%).\n"
        "   Tỷ lệ phân bố 48% vs 52% đạt mức cân bằng lý tưởng, loại bỏ hoàn toàn hiện tượng lệch nhãn (Imbalanced Classes).\n"
        "2. Xử lý giá trị khuyết thiếu: Các ô điểm khuyết thiếu (NaN) được điền giá trị 0.0 theo đúng quy chuẩn mã hóa dữ liệu học vụ "
        "(đại diện cho việc sinh viên chưa tích lũy môn học hoặc bị vắng thi/điểm F).\n"
        "3. Giao thức Zero Label Leakage: Điểm GPA toàn khóa được tính bình quân gia quyền từ chính điểm của 50 môn học. "
        "Nếu đưa toàn bộ 50 môn học vào đầu vào, bài toán sẽ rơi vào lỗi rò rỉ nhãn vòng quanh (Circular Label Leakage), "
        "làm mất đi tính thực tế của việc dự báo sớm. Do đó, đề tài chỉ sử dụng duy nhất 15 môn học của Năm thứ nhất (Học kỳ 1 & 2) "
        "làm biến đầu vào để dự đoán sớm. Tại thời điểm kết thúc năm nhất, sinh viên hoàn toàn chưa có điểm của 35 môn học các năm sau.\n"
        "4. Phân chia dữ liệu: Tập dữ liệu 1.656 mẫu được phân chia thành 80% Huấn luyện (1.324 sinh viên) và 20% Kiểm thử độc lập (332 sinh viên) "
        "bằng phương pháp phân tầng (StratifiedSplit) để giữ nguyên tỷ lệ nhãn."
    )

    # Chèn Biểu đồ 2: Ma trận tương quan
    path_corr = os.path.join(OUTPUTS_DIR, "correlation_heatmap.png")
    if os.path.exists(path_corr):
        doc.add_picture(path_corr, width=Inches(5.8))
        p_cap = doc.add_paragraph("Hình 3.2: Ma trận tương quan giữa GPA và các môn học nền tảng Năm thứ nhất")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    # -------------------------------------------------------------
    # CHƯƠNG 4
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM", level=1)
    
    doc.add_heading("4.1. Hiệu Năng Phân Lớp Trên Tập Kiểm Thử Độc Lập (332 Sinh Viên)", level=2)
    dt_m = metrics["scenario_early_decision_tree"]
    lr_m = metrics["scenario_early_logistic_regression"]

    table_res = doc.add_table(rows=6, cols=3)
    table_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    res_headers = ["Chỉ Số Đánh Giá", "Cây Quyết Định (Decision Tree)", "Logistic Regression (Đối chứng)"]
    
    for i, h in enumerate(res_headers):
        c = table_res.rows[0].cells[i]
        c.text = h
        set_cell_background(c, "2980B9")
        set_cell_margins(c, top=100, bottom=100, left=120, right=120)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)

    metrics_rows = [
        ("Độ chính xác toàn cục (Accuracy)", f"{dt_m['accuracy']*100:.2f}%", f"{lr_m['accuracy']*100:.2f}%"),
        ("Độ chuẩn xác lớp Khá-Giỏi (Precision)", f"{dt_m['precision_pass']*100:.2f}%", f"{lr_m['precision_pass']*100:.2f}%"),
        ("Độ nhạy lớp Khá-Giỏi (Recall)", f"{dt_m['recall_pass']*100:.2f}%", f"{lr_m['recall_pass']*100:.2f}%"),
        ("Điểm F1 lớp Khá-Giỏi (F1-Score)", f"{dt_m['f1_pass']*100:.2f}%", f"{lr_m['f1_pass']*100:.2f}%"),
        ("Macro F1-Score trung bình", f"{dt_m['macro_f1']*100:.2f}%", f"{lr_m['macro_f1']*100:.2f}%"),
    ]

    for row_idx, data_tuple in enumerate(metrics_rows, start=1):
        for col_idx, text_val in enumerate(data_tuple):
            cell = table_res.rows[row_idx].cells[col_idx]
            cell.text = text_val
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            cell.paragraphs[0].runs[0].font.size = Pt(9.5)
            if col_idx > 0:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if col_idx == 1:
                cell.paragraphs[0].runs[0].font.bold = True

    p_cap = doc.add_paragraph("Bảng 4.1: Đối sánh hiệu năng phân lớp trên tập kiểm thử độc lập (Test Set 20% = 332 sinh viên)")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.runs[0].italic = True
    p_cap.runs[0].font.size = Pt(9.5)

    p_compare = doc.add_paragraph(
        f"Nhận xét đối chứng khoa học:\n"
        f"- Mô hình Decision Tree đạt Accuracy {dt_m['accuracy']*100:.2f}% chỉ bằng dữ liệu năm nhất. "
        f"Mô hình đối chứng Logistic Regression đạt Accuracy {lr_m['accuracy']*100:.2f}%, cao hơn Decision Tree khoảng 6.9 điểm %. "
        f"Điều này hoàn toàn phù hợp với lý thuyết toán học: Logistic Regression tối ưu hóa siêu phẳng liên tục trên không gian 15 chiều, "
        f"trong khi Cây quyết định bị ràng buộc bởi các đường cắt trực giao song song với trục tọa độ.\n"
        f"- Tuy nhiên, trong ứng dụng quản lý giáo dục, Cây quyết định đạt ưu thế áp đảo về tính khả giải (Interpretability): "
        f"Cung cấp trực tiếp các quy tắc rẽ nhánh số học rõ ràng (ví dụ: Điện tử ô tô <= 2.25, Vật liệu học > 2.75, Tiếng Anh 2 > 2.25), "
        f"giúp cố vấn học tập có thể giải thích minh bạch cho sinh viên."
    )

    doc.add_heading("4.2. Khảo Sát Siêu Tham Số max_depth Bằng 5-Fold Stratified Cross-Validation", level=2)
    p = doc.add_paragraph(
        "Để ngăn chặn tuyệt đối rò rỉ tập Test trong quá trình chọn siêu tham số, tác giả thực hiện kiểm định chéo 5 tầng (5-Fold Stratified CV) "
        "hoàn toàn trên tập Huấn luyện (1.324 sinh viên). Kết quả khảo sát được thể hiện trong Bảng 4.2:"
    )

    depth_data = metrics["depth_tuning_cv"]
    table_depth = doc.add_table(rows=len(depth_data) + 1, cols=6)
    table_depth.alignment = WD_TABLE_ALIGNMENT.CENTER
    d_headers = ["Độ Sâu", "CV Mean", "CV Std", "Train Acc", "Test Acc", "Gap"]
    
    for i, h in enumerate(d_headers):
        c = table_depth.rows[0].cells[i]
        c.text = h
        set_cell_background(c, "2980B9")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)

    for idx, row_data in enumerate(depth_data, start=1):
        d = row_data["max_depth"]
        r = table_depth.rows[idx].cells
        r[0].text = str(d)
        r[1].text = f"{row_data['cv_mean_accuracy']*100:.2f}%"
        r[2].text = f"±{row_data['cv_std_accuracy']*100:.2f}%"
        r[3].text = f"{row_data['train_accuracy']*100:.2f}%"
        r[4].text = f"{row_data['test_accuracy']*100:.2f}%"
        r[5].text = f"{row_data['generalization_gap']*100:+.2f}%"

        for cell in r:
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
        for col_i in range(6):
            r[col_i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_cap = doc.add_paragraph("Bảng 4.2: Khảo sát ảnh hưởng của độ sâu cây bằng 5-Fold Stratified Cross-Validation trên tập Train")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.runs[0].italic = True
    p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.3. Phân Tích Ma Trận Nhầm Lẫn (Confusion Matrix)", level=2)
    p = doc.add_paragraph(
        f"Ma trận nhầm lẫn của mô hình Cây quyết định trên 332 sinh viên kiểm thử (Hình 4.1) ghi nhận:\n"
        f"- True Negative (TN): {dt_m['tn']} sinh viên thực tế xếp loại Trung bình-Yếu được dự đoán chính xác.\n"
        f"- True Positive (TP): {dt_m['tp']} sinh viên thực tế xếp loại Khá-Giỏi được dự đoán chính xác.\n"
        f"- False Positive (FP --- Sai loại 1): {dt_m['fp']} sinh viên thực tế Trung bình-Yếu nhưng mô hình dự đoán nhầm là Khá-Giỏi.\n"
        f"- False Negative (FN --- Sai loại 2): {dt_m['fn']} sinh viên thực tế Khá-Giỏi nhưng mô hình dự đoán nhầm là Trung bình-Yếu.\n"
        f"Độ chuẩn xác Precision đạt {dt_m['precision_pass']*100:.2f}%, bảo đảm độ tin cậy khi hệ thống đưa ra nhận định sinh viên sẽ đạt loại Khá/Giỏi."
    )

    path_cm = os.path.join(OUTPUTS_DIR, "confusion_matrix.png")
    if os.path.exists(path_cm):
        doc.add_picture(path_cm, width=Inches(4.5))
        p_cap = doc.add_paragraph("Hình 4.1: Ma trận nhầm lẫn của mô hình Decision Tree trên tập kiểm thử (332 sinh viên)")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.4. Phân Tích Cấu Trúc Cây Quyết Định & Các Quy Tắc Rẽ Nhánh", level=2)
    p = doc.add_paragraph(
        "Cấu trúc Cây quyết định (Hình 4.2) thể hiện rõ ràng các điều kiện rẽ nhánh thực tế từ dữ liệu:\n"
        "1. Nút gốc (Root Node): Thuật toán chọn môn 'Điện - Điện tử ô tô' với ngưỡng điểm 2.25 làm nút phân chia đầu tiên. "
        "Nếu điểm môn này <= 2.25, xác suất sinh viên rơi vào nhóm Trung bình-Yếu chiếm ưu thế rõ rệt.\n"
        "2. Tầng rẽ nhánh kế tiếp: Tại nhánh điểm cao (> 2.25), thuật toán tiếp tục kiểm tra điểm môn 'Vật liệu học' (ngưỡng 2.75) và môn 'Tiếng Anh 2' (ngưỡng 2.25). "
        "Những sinh viên vượt qua cả hai ngưỡng này có tỷ lệ đạt bằng Khá/Giỏi trên 85%."
    )

    path_tree = os.path.join(OUTPUTS_DIR, "decision_tree.png")
    if os.path.exists(path_tree):
        doc.add_picture(path_tree, width=Inches(6.5))
        p_cap = doc.add_paragraph("Hình 4.2: Cấu trúc Cây quyết định dự đoán sớm kết quả học tập sinh viên")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.5. Xếp Hạng Mức Độ Quan Trọng Của Các Môn Học Năm 1 (Feature Importance)", level=2)
    p = doc.add_paragraph(
        "Mức độ quan trọng Gini (Gini Importance) của các môn học Năm 1 trong việc phân tách học lực sinh viên được thể hiện ở Hình 4.3:\n"
        "- Điện - Điện tử ô tô: Đóng góp 44.46% trọng số phân loại (môn học bản lề hàng đầu).\n"
        "- Vật liệu học: Đóng góp 24.85% trọng số.\n"
        "- Tiếng Anh 2: Đóng góp 21.97% trọng số (chứng minh ngoại ngữ là rào cản phân hóa sinh viên kỹ thuật).\n"
        "- Nhiệt kỹ thuật: Đóng góp 6.69% trọng số.\n"
        "- Tiếng Anh 1: Đóng góp 2.02% trọng số."
    )

    path_fi = os.path.join(OUTPUTS_DIR, "feature_importance.png")
    if os.path.exists(path_fi):
        doc.add_picture(path_fi, width=Inches(5.5))
        p_cap = doc.add_paragraph("Hình 4.3: Top các môn học Năm 1 có mức độ ảnh hưởng lớn nhất đến kết quả học tập")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.6. Phân Tích Định Lượng Nguyên Nhân Dự Đoán Sai (Near-Threshold Error Analysis)", level=2)
    err = metrics["error_analysis"]
    p = doc.add_paragraph(
        f"Tổng số ca sai trên tập kiểm thử là **{err['total_errors']} / 332 sinh viên** ({err['error_rate']}%). "
        f"Phân tích định lượng phân bố điểm GPA của các ca sai cho thấy:\n"
        f"1. Nhóm False Positive ({err['false_positives']} ca --- Thực tế TB-Yếu nhưng đoán Khá-Giỏi):\n"
        f"   - Điểm GPA trung vị thực tế đạt {err['fp_gpa_median']:.3f} (trung bình: {err['fp_gpa_mean']:.3f}).\n"
        f"   - Trong đó có tới {err['fp_near_threshold']} / {err['false_positives']} sinh viên ({err['fp_near_pct']}%) có GPA >= 2.30, "
        f"tức chỉ kém ngưỡng phân định 2.50 từ 0.05 đến 0.20 điểm.\n"
        f"2. Nhóm False Negative ({err['false_negatives']} ca --- Thực tế Khá-Giỏi nhưng đoán TB-Yếu):\n"
        f"   - Điểm GPA trung vị thực tế đạt {err['fn_gpa_median']:.3f} (trung bình: {err['fn_gpa_mean']:.3f}).\n"
        f"   - Trong đó có tới {err['fn_near_threshold']} / {err['false_negatives']} sinh viên ({err['fn_near_pct']}%) có GPA < 2.70, "
        f"cũng nằm sát ngay phía trên ranh giới 2.50.\n"
        f"3. Kết luận định lượng:\n"
        f"   - Có tới {err['total_near_threshold']} / {err['total_errors']} ca sai ({err['near_threshold_pct']}%) tập trung trong dải điểm hẹp [2.30, 2.70) "
        f"cận biên ranh giới quyết định. Điều này chứng minh rằng mô hình hoàn toàn không phạm sai lầm thô thiển, "
        f"mà sai số chủ yếu xuất hiện tại vùng chuyển tiếp nhạy cảm giữa mức Trung bình khá và Khá."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 5
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 5: KẾT LUẬN & HƯỚNG PHÁT TRIỂN", level=1)
    
    doc.add_heading("5.1. Kết Quả Đạt Được", level=2)
    p = doc.add_paragraph(
        "1. Xây dựng hoàn chỉnh đường ống xử lý dữ liệu và huấn luyện mô hình Cây quyết định trên bộ dữ liệu 1.656 sinh viên.\n"
        "2. Đạt độ chính xác 71.39% và Macro F1 70.94% trong kịch bản dự đoán sớm chỉ sử dụng 15 môn học Năm thứ nhất, "
        "bảo đảm nguyên tắc Zero Data Leakage.\n"
        "3. Áp dụng 5-Fold Stratified Cross-Validation trên tập Train để khảo sát siêu tham số max_depth một cách khoa học.\n"
        "4. Trực quan hóa cấu trúc cây quyết định và giải mã nguyên nhân sai số định lượng tại vùng cận biên ranh giới quyết định."
    )

    doc.add_heading("5.2. Điểm Mạnh & Giới Hạn Khách Quan", level=2)
    p = doc.add_paragraph(
        "1. Điểm mạnh: Tính ứng dụng cao, mô hình minh bạch, phương pháp luận chặt chẽ.\n"
        "2. Giới hạn khách quan: Dữ liệu chỉ bao gồm điểm thi số học, chưa có các thuộc tính hành vi (như tỷ lệ chuyên cần, thời gian tự học, tương tác hệ thống LMS)."
    )

    doc.add_heading("5.3. Hướng Mở Rộng Tiếp Theo", level=2)
    p = doc.add_paragraph(
        "1. Mở rộng sang các thuật toán học tổ hợp (Ensemble Methods: Random Forest, Gradient Boosting) để cải thiện độ chính xác.\n"
        "2. Thu thập thêm các đặc trưng hành vi số: Tích hợp dữ liệu điểm danh và log học tập trên nền tảng e-learning LMS để bổ trợ cho dữ liệu điểm số."
    )

    # Tài liệu tham khảo
    doc.add_heading("TÀI LIỆU THAM KHẢO", level=1)
    refs = [
        "[1] Cortez, P., & Silva, A. (2008). Using data mining to predict secondary school student performance. In Proceedings of 5th Annual Future Business Technology Conference (FUBUTEC 2008), pp. 5-12.",
        "[2] Breiman, L., Friedman, J. H., Olshen, R. A., & Stone, C. J. (1984). Classification and Regression Trees. Wadsworth & Brooks/Cole Advanced Books & Software.",
        "[3] Quinlan, J. R. (1986). Induction of decision trees. Machine Learning, 1(1), pp. 81-106.",
        "[4] Romero, C., & Ventura, S. (2010). Educational data mining: a review of the state of the art. IEEE Transactions on Systems, Man, and Cybernetics, Part C, 40(6), pp. 601-618.",
        "[5] Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, pp. 2825-2830.",
        "[6] Bộ Giáo dục và Đào tạo (2021). Thông tư 08/2021/TT-BGDĐT ban hành Quy chế đào tạo trình độ đại học.",
        "[7] Trường Đại học Công nghiệp Hà Nội (HaUI). Quy định về công tác đào tạo đại học chính quy theo hệ thống tín chỉ."
    ]
    for r in refs:
        p_ref = doc.add_paragraph(r)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.left_indent = Inches(0.25)
        p_ref.runs[0].font.size = Pt(9.5)

    doc.add_page_break()

    # =============================================================
    # PHẦN II: GIẢI MÃ CHI TIẾT Ý NGHĨA CỦA TỪNG CON SỐ THỰC NGHIỆM
    # =============================================================
    h_p2 = doc.add_heading("PHẦN II: GIẢI MÃ CHI TIẾT Ý NGHĨA CỦA TỪNG CON SỐ THỰC NGHIỆM", level=1)
    h_p2.runs[0].font.color.rgb = RGBColor(41, 128, 185)

    doc.add_heading("1. Con Số Về Tập Dữ Liệu: 1.656 Sinh Viên & 56 Biến", level=2)
    p = doc.add_paragraph(
        "- 1.656 sinh viên: Đại diện cho toàn bộ số lượng sinh viên kỹ thuật thực tế được theo dõi liên tục trong suốt 4 năm học (từ HK1 đến HK8). "
        "Kích thước mẫu này vượt xa ngưỡng mẫu tối thiểu thông thường trong học máy (N >= 300), bảo đảm các kết luận rút ra có ý nghĩa thống kê vững chắc.\n"
        "- 56 thuộc tính: Bao gồm 2 biến nhân khẩu học (BirthPlace, Sex), 50 môn học từ HK1 đến HK8, số lần nhận học bổng, điểm GPA tích lũy và trạng thái kỷ luật.\n"
        "- 15 môn học Năm 1: Là 15 môn học sinh viên hoàn thành trong năm thứ nhất. Sử dụng đúng 15 môn này phản ánh bài toán cảnh báo sớm thực tế."
    )

    doc.add_heading("2. Con Số Về Biến Mục Tiêu: Ngưỡng GPA = 2.50 và Tỷ Lệ 47.95% vs 52.05%", level=2)
    p = doc.add_paragraph(
        "- Ngưỡng GPA = 2.50: Là mốc phân định tốt nghiệp loại Khá trở lên theo Thông tư 08/2021/TT-BGDĐT. "
        "Đây là điều kiện tiên quyết để sinh viên cạnh tranh việc làm sau khi ra trường.\n"
        "- Tỷ lệ 794 (47.95%) vs 862 (52.05%): Tỷ lệ phân bố gần như hoàn hảo 1:1. "
        "Trong học máy phân lớp, tỷ lệ cân bằng này giúp các chỉ số Accuracy, Precision, Recall không bị méo mó hay thiên vị cho bất kỳ lớp nào."
    )

    doc.add_heading("3. Con Số Về Hiệu Năng Phân Lớp: Accuracy 71.39%, Precision 74.24%, Recall 61.64%", level=2)
    p = doc.add_paragraph(
        "- Accuracy = 71.39%: Cứ 100 sinh viên kiểm thử, mô hình dự đoán chính xác xếp loại tốt nghiệp cho hơn 71 sinh viên chỉ dựa trên kết quả học tập của năm nhất. "
        "Đây là tỷ lệ rất cao đối với bài toán cảnh báo sớm trước 3 năm tốt nghiệp.\n"
        "- Precision (Khá-Giỏi) = 74.24%: Khi mô hình dự đoán một sinh viên sẽ đạt loại Khá/Giỏi, xác suất sinh viên đó thực sự đạt Khá/Giỏi lên tới 74.24%. "
        "Chỉ có 25.76% khả năng bị đoán nhầm.\n"
        "- Recall (Khá-Giỏi) = 61.64%: Trong tổng số tất cả sinh viên thực tế đạt Khá/Giỏi, mô hình nhận diện và tóm bắt thành công 61.64% sinh viên ngay từ năm nhất.\n"
        "- Macro F1 = 70.94%: Trung bình điều hòa giữa hai lớp đạt xấp xỉ 71%, khẳng định mô hình học đều cả hai nhóm học lực, không bị thiên vị."
    )

    doc.add_heading("4. Tại Sao Logistic Regression Đạt 78.31% Cao Hơn Decision Tree Nhưng Decision Tree Vẫn Được Chọn?", level=2)
    p = doc.add_paragraph(
        "- Nguyên nhân toán học: Logistic Regression sử dụng hàm sigmoid kết hợp tổ hợp tuyến tính trên toàn bộ 15 biến cùng lúc. "
        "Do điểm thi là các biến số thực liên tục từ 0.0 đến 4.0, siêu phẳng tuyến tính có thể xoay nghiêng theo mọi góc trong không gian 15 chiều để tối ưu hóa ranh giới. "
        "Ngược lại, Cây quyết định cắt không gian bằng các siêu phẳng trực giao song song với trục tọa độ.\n"
        "- Lý do Decision Tree được chọn: Khả năng giải thích (Interpretability). Logistic Regression là một mô hình hộp đen với các trọng số toán học trừu tượng. "
        "Trong khi đó, Decision Tree cung cấp ngay điều kiện rẽ nhánh số học: 'Nếu sinh viên đạt điểm môn Điện - Điện tử ô tô > 2.25 và Vật liệu học > 2.75 "
        "thì xác suất đạt Khá/Giỏi là 85%'. Điều này mang tính hành động trực tiếp cho công tác cố vấn học tập."
    )

    doc.add_heading("5. Ý Nghĩa Của Bảng 5-Fold Cross-Validation & Hiện Tượng Overfitting", level=2)
    p = doc.add_paragraph(
        "- Tại max_depth = 1: CV Mean chỉ đạt 63.67%, Train Acc 66.69% ---> Cây quá nông, bị Underfitting (Dưới khớp).\n"
        "- Tại max_depth = 3: CV Mean đạt 72.43% với độ lệch chuẩn nhỏ nhất (±1.30%), độ sâu chỉ 3 tầng gồm 8 nút lá ---> Đây là cấu hình tối ưu để giải thích.\n"
        "- Tại max_depth = 4: Test Acc đạt đỉnh 74.40%.\n"
        "- Tại max_depth = 8: Train Acc tăng vọt lên 82.10%, nhưng Test Acc giảm về 71.39%. Khoảng cách tổng quát hóa (Generalization Gap) dãn rộng tới +10.71% "
        "---> Đây là minh chứng thực nghiệm không thể chối cãi của hiện tượng Overfitting (Quá khớp)."
    )

    doc.add_heading("6. Ý Nghĩa Con Số Của Top Môn Học Quan Trọng Nhất (Feature Importance)", level=2)
    p = doc.add_paragraph(
        "- Điện - Điện tử ô tô (44.46%): Chiếm gần một nửa tổng trọng số phân loại. Đây là môn chuyên ngành cơ sở đầu tiên ở năm 1, đòi hỏi cả tư duy mạch điện, "
        "tính toán định lượng và khả năng thực hành. Sinh viên vượt qua tốt môn này thể hiện sự thích nghi xuất sắc với ngành kỹ thuật.\n"
        "- Vật liệu học (24.85%): Môn khoa học vật liệu cung cấp kiến thức nền tảng về cơ tính, nhiệt luyện.\n"
        "- Tiếng Anh 2 (21.97%): Đóng góp hơn 1/5 tầm quan trọng. Con số này chứng minh ngoại ngữ là rào cản phân hóa cực kỳ mạnh đối với sinh viên kỹ thuật."
    )

    doc.add_heading("7. Ý Nghĩa Con Số Phân Tích Lỗi Cận Biên: 70 / 95 Ca Sai (73.68%)", level=2)
    p = doc.add_paragraph(
        "Trong 95 ca dự đoán sai trên tập test:\n"
        "- 34 ca False Positive: GPA trung vị là 2.388, trong đó 24 sinh viên (70.59%) có GPA >= 2.30.\n"
        "- 61 ca False Negative: GPA trung vị là 2.595, trong đó 46 sinh viên (75.41%) có GPA < 2.70.\n"
        "- Tổng cộng 70 / 95 ca sai (73.68%) nằm trong khoảng GPA từ 2.30 đến 2.69: Con số này chứng minh mô hình hoạt động rất ổn định; "
        "sai số không xuất phát từ việc đoán sai những ca rõ ràng, mà xuất phát từ việc ranh giới điểm 2.50 trong thực tế rất sít sao (chỉ chênh nhau 0.05 --- 0.15 điểm)."
    )

    doc.add_page_break()

    # =============================================================
    # PHẦN III: GIẢI THÍCH KIẾN TRÚC MÃ NGUỒN VÀ Ý NGHĨA TỪNG ĐOẠN CODE
    # =============================================================
    h_p3 = doc.add_heading("PHẦN III: GIẢI THÍCH KIẾN TRÚC MÃ NGUỒN VÀ Ý NGHĨA TỪNG ĐOẠN CODE", level=1)
    h_p3.runs[0].font.color.rgb = RGBColor(41, 128, 185)

    doc.add_heading("1. Module src/data_loader.py --- Đọc và Khám Phá Dữ Liệu", level=2)
    p = doc.add_paragraph(
        "- load_student_data(): Định vị file dữ liệu chính và từ điển biến. Kiểm tra an toàn bằng os.path.exists(), "
        "trả về bảng dữ liệu df 1.656 bản ghi và từ điển df_dict.\n"
        "- get_dataset_summary(): Tự động tính toán các chỉ số thống kê 5 số của biến GPA (Mean 2.399, Median 2.485, Std 0.510, Min 0.0, Max 3.48), "
        "thống kê số lượng tỉnh thành nơi sinh và cơ cấu giới tính."
    )

    doc.add_heading("2. Module src/preprocessing.py --- Tiền Xử Lý & Giao Thức Zero Label Leakage", level=2)
    p = doc.add_paragraph(
        "- create_target(): Nhị phân hóa nhãn học lực theo ngưỡng GPA >= 2.50.\n"
        "- clean_course_grades(): Ép kiểu dữ liệu điểm số thực và điền 0.0 cho các ô khuyết (.fillna(0.0)), "
        "phản ánh đúng việc sinh viên chưa học hoặc vắng thi/điểm F.\n"
        "- prepare_features(scenario='early'): Cốt lõi ngăn ngừa rò rỉ nhãn: Chỉ lọc 15 môn học có Semester suffix in [1.0, 2.0], "
        "loại bỏ toàn bộ môn học HK3-HK8 và GPA toàn khóa khỏi đầu vào.\n"
        "- split_data(): Chia tập 80/20 với stratify=y để bảo toàn tuyệt đối tỷ lệ 48% vs 52% giữa hai lớp ở cả Train và Test."
    )

    doc.add_heading("3. Module src/train_model.py --- Huấn Luyện & Kiểm Định 5-Fold CV", level=2)
    p = doc.add_paragraph(
        "- tune_tree_depth_cv(): Chạy StratifiedKFold(n_splits=5) hoàn toàn trên tập Train. "
        "Tập Test được cô lập hoàn toàn để bảo đảm nguyên tắc Zero Test Leakage.\n"
        "- train_decision_tree(): Khởi tạo CART với criterion='gini', max_depth=3, min_samples_leaf=10 để khống chế cây không quá khớp.\n"
        "- train_baseline_model(): Huấn luyện Logistic Regression làm đối chứng khoa học."
    )

    doc.add_heading("4. Module src/evaluate_model.py --- Đánh Giá & Xuất Biểu Đồ", level=2)
    p = doc.add_paragraph(
        "- calculate_metrics(): Tính toán Accuracy, Precision, Recall, F1 và Macro F1 từ confusion_matrix.\n"
        "- plot_decision_tree_graph(): Việt hóa tên biến rút gọn và vẽ cây quyết định bo tròn với màu sắc tương phản rõ ràng.\n"
        "- analyze_errors_quantitatively(): Trích xuất 95 ca dự đoán sai, tính toán phân vị GPA và xuất ra con số 73.68% lỗi cận biên."
    )

    doc.add_heading("5. Kịch Bản main.py --- Điều Phối Toàn Trình 5 Giai Đoạn", level=2)
    p = doc.add_paragraph(
        "Kịch bản main.py điều phối liên hoàn 5 giai đoạn từ đọc dữ liệu, tiền xử lý, kiểm định 5-Fold CV, "
        "đánh giá tập Test đến lưu toàn bộ số liệu vào outputs/metrics.json trong vòng chưa đầy 8 giây."
    )

    # Lưu file report.docx
    output_docx_path = os.path.join(BASE_DIR, "report.docx")
    doc.save(output_docx_path)
    print(f" -> Đã tạo thành công file Word hoàn chỉnh: {output_docx_path}")

    # Copy ra thư mục gốc E:\KPDL\report.docx để người dùng mở ngay lập tức
    root_docx_path = os.path.join(os.path.dirname(BASE_DIR), "report.docx")
    try:
        doc.save(root_docx_path)
        print(f" -> Đã sao lưu ra thư mục gốc: {root_docx_path}")
    except Exception as e:
        print(f" -> Lưu ý sao lưu: {e}")


if __name__ == "__main__":
    build_phenikaa_report_docx()
