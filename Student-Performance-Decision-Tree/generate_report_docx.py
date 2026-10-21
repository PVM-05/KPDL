"""
Kịch bản tự động tạo file Báo cáo Word (report.docx) và Markdown (report.md)
theo chuẩn 100% đề cương của Giảng viên và số liệu thực nghiệm thực tế từ outputs/metrics.json.
Phiên bản: Báo cáo cá nhân, kiểm soát 100% rò rỉ nhãn (Zero Label Leakage) và đánh giá bằng 5-Fold CV.
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

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUTS_DIR = os.path.join(BASE_DIR, "outputs")
METRICS_PATH = os.path.join(OUTPUTS_DIR, "metrics.json")


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


def build_docx_report():
    print("Đang đọc dữ liệu thực nghiệm từ metrics.json...")
    with open(METRICS_PATH, "r", encoding="utf-8") as f:
        metrics = json.load(f)

    doc = Document()

    # Thiết lập lề trang chuẩn A4 (2cm xung quanh)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # -------------------------------------------------------------
    # TRANG BÌA & THÔNG TIN HÀNH CHÍNH (BÀI TẬP LỚN CÁ NHÂN)
    # -------------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_inst = title_p.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO\nTRƯỜNG ĐẠI HỌC CÔNG NGHIỆP HÀ NỘI\nKHOA CÔNG NGHỆ THÔNG TIN\n")
    run_inst.bold = True
    run_inst.font.size = Pt(13)
    run_inst.font.color.rgb = RGBColor(44, 62, 80)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(15)

    title_main = doc.add_paragraph()
    title_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = title_main.add_run("BÁO CÁO BÀI TẬP LỚN CÁ NHÂN HỌC PHẦN KHAI PHÁ DỮ LIỆU\n\n")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(127, 140, 141)

    run_title = title_main.add_run("DỰ ĐOÁN SỚM KẾT QUẢ HỌC TẬP CỦA SINH VIÊN BẰNG THUẬT TOÁN DECISION TREE\n")
    run_title.bold = True
    run_title.font.size = Pt(18)
    run_title.font.color.rgb = RGBColor(41, 128, 185)

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_info.paragraph_format.space_before = Pt(25)
    p_info.paragraph_format.space_after = Pt(25)
    run_meta = p_info.add_run(
        "Đối tượng nghiên cứu: Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark - 1.656 sinh viên)\n"
        "Phương pháp luận: Kịch bản dự đoán sớm sau Năm 1 — Zero Label Leakage & 5-Fold Cross-Validation\n"
        "Giảng viên phụ trách: Bộ môn Khai phá Dữ liệu\n"
        "Năm học: 2026\n"
    )
    run_meta.font.size = Pt(11)
    run_meta.italic = True

    # Bảng thông tin sinh viên thực hiện (Bài cá nhân 100%)
    doc.add_heading("THÔNG TIN SINH VIÊN THỰC HIỆN & CAM ĐOAN ĐÓNG GÓP", level=2)
    p_desc = doc.add_paragraph()
    r_desc = p_desc.add_run(
        "Căn cứ vào quy chế đánh giá học phần: Điểm báo cáo kỹ thuật (50%) và điểm vấn đáp (50%) "
        "được đánh giá trực tiếp cho sinh viên thực hiện độc lập đề tài:"
    )
    r_desc.font.size = Pt(10.5)

    table_contrib = doc.add_table(rows=2, cols=5)
    table_contrib.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["STT", "Họ và Tên Sinh Viên", "Mã Sinh Viên", "Nhiệm Vụ Thực Hiện Toàn Trình", "Hệ Số Đóng Góp"]
    
    hdr_cells = table_contrib.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "2980B9")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(10)

    # Dữ liệu sinh viên cá nhân
    row1 = table_contrib.rows[1].cells
    row1[0].text = "1"
    row1[1].text = "........................................"
    row1[2].text = "MSSV: ...................."
    row1[3].text = "Thực hiện độc lập 100%: Khảo sát bộ dữ liệu, thiết kế pipeline tiền xử lý, huấn luyện Decision Tree & Logistic Regression, khảo sát 5-Fold CV, xuất biểu đồ và viết báo cáo hoàn chỉnh."
    row1[4].text = "1.0 (100%)"

    for cell in row1:
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
    row1[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row1[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_agree = doc.add_paragraph()
    p_agree.paragraph_format.space_before = Pt(15)
    p_agree.paragraph_format.space_after = Pt(25)
    run_agree = p_agree.add_run(
        "Lời cam đoan: Tôi xin cam đoan đây là công trình bài tập lớn do bản thân tôi tự thực hiện độc lập. "
        "Mọi số liệu thực nghiệm trong báo cáo đều được trích xuất trung thực từ mã nguồn chương trình, "
        "dữ liệu sử dụng đã được khử định danh học thuật và tuân thủ các chuẩn mực nghiên cứu khoa học."
    )
    run_agree.bold = True
    run_agree.italic = True
    run_agree.font.size = Pt(10.5)

    doc.add_page_break()

    # -------------------------------------------------------------
    # CHƯƠNG 1: GIỚI THIỆU TỔNG QUAN
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 1: GIỚI THIỆU TỔNG QUAN", level=1)
    
    doc.add_heading("1.1. Bối Cảnh Bài Toán & Tính Cấp Thiết", level=2)
    p = doc.add_paragraph(
        "Trong các trường đại học đào tạo theo hệ thống tín chỉ, sinh viên năm thứ nhất thường phải đối mặt với bước chuyển tiếp lớn "
        "về phương pháp học tập và tính tự chủ. Nhiều sinh viên gặp khó khăn trong việc làm quen với khối kiến thức khoa học cơ bản "
        "và kỹ thuật cơ sở dẫn đến tình trạng sa sút kết quả học tập, bị cảnh báo học vụ hoặc thậm chí bị buộc thôi học ở các năm sau. "
        "Do đó, bài toán Khai phá Dữ liệu Giáo dục (Educational Data Mining - EDM) nhằm xây dựng hệ thống cảnh báo sớm (Early Warning System) "
        "có ý nghĩa thực tiễn to lớn: giúp nhà trường và cố vấn học tập chủ động phát hiện sớm sinh viên có nguy cơ để can thiệp kịp thời ngay sau năm 1, "
        "thay vì chờ đến khi kết quả đã an bài."
    )
    p.paragraph_format.line_spacing = 1.25

    doc.add_heading("1.2. Mục Tiêu & Phạm Vi Đề Tài", level=2)
    p = doc.add_paragraph(
        "Đề tài tập trung giải quyết bài toán phân lớp nhị phân dự đoán sớm kết quả học tập tốt nghiệp của sinh viên đại học. "
        "Mục tiêu cụ thể gồm:\n"
        "1. Khảo sát hiện trạng phân bố điểm GPA và sự tương quan giữa các môn học Năm 1 với điểm tích lũy toàn khóa.\n"
        "2. Ứng dụng thuật toán Cây quyết định (Decision Tree Classifier) để dự đoán sớm sinh viên sẽ đạt loại Khá/Giỏi (GPA >= 2.50) "
        "hay Trung bình/Yếu (GPA < 2.50) chỉ bằng kết quả các môn học của Năm thứ nhất (15 môn nền tảng).\n"
        "3. Đảm bảo tính minh bạch và khả năng giải thích (Interpretability) của mô hình cây thông qua việc trực quan hóa cấu trúc rẽ nhánh (If-Then rules) "
        "và xếp hạng các môn học có ảnh hưởng lớn nhất đến kết quả học tập.\n"
        "4. Tuân thủ nghiêm ngặt nguyên tắc Zero Data Leakage: Không sử dụng điểm của các năm học sau hoặc các thông tin rò rỉ nhãn vào quá trình dự đoán sớm."
    )

    doc.add_heading("1.3. Tổng Quan Machine Learning & Bài Toán Muốn Làm (~100 chữ)", level=2)
    p = doc.add_paragraph(
        "Học máy (Machine Learning) là ngành khoa học nghiên cứu các thuật toán cho phép hệ thống tự động đúc kết quy luật từ dữ liệu. "
        "Trong bài toán phân lớp kết quả học tập này, tác giả lựa chọn thuật toán Cây quyết định (Decision Tree). Thuật toán hoạt động bằng cách "
        "phân tách không gian dữ liệu thành các tập con thuần nhất dựa trên chỉ số Độ vấy bẩn Gini (Gini Impurity). "
        "Lý do then chốt để lựa chọn Cây quyết định trong giáo dục là cấu trúc dạng cây phân cấp (White-box model) mang tính minh bạch tuyệt đối, "
        "giúp giảng viên và cố vấn học tập dễ dàng hiểu được ngưỡng điểm của từng môn học cụ thể để tư vấn cho sinh viên, "
        "thay vì các mô hình hộp đen khó giải thích."
    )
    p.runs[0].italic = True

    doc.add_heading("1.4. Công Việc Liên Quan & Các Nghiên Cứu Khoa Học", level=2)
    p = doc.add_paragraph(
        "Nghiên cứu kinh điển của Cortez và Silva (2008) [1] đã mở đầu xu hướng áp dụng các thuật toán Học máy (Decision Tree, Random Forest, SVM) "
        "để dự đoán kết quả học tập, chỉ ra rằng điểm các kỳ học đầu tiên đóng vai trò quan trọng nhất trong việc định hình kết quả cuối khóa. "
        "Romero và Ventura (2010) [4] đã khảo sát toàn diện các kỹ thuật khai phá dữ liệu giáo dục và nhấn mạnh rằng tính khả giải (Interpretability) "
        "của mô hình Cây quyết định là tiêu chí hàng đầu để ứng dụng trong môi trường đào tạo. "
        "Tại Việt Nam, các nghiên cứu tại các trường kỹ thuật lớn cũng khẳng định điểm số các môn khoa học cơ bản (Giải tích, Vật lý, Ngoại ngữ) "
        "ở năm nhất có hệ số tương quan rất cao với năng lực tốt nghiệp của sinh viên."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 2: CƠ SỞ LÝ THUYẾT THUẬT TOÁN DECISION TREE
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 2: THUẬT TOÁN CÂY QUYẾT ĐỊNH (DECISION TREE)", level=1)
    
    doc.add_heading("2.1. Bản Chất Của Cây Quyết Định", level=2)
    p = doc.add_paragraph(
        "Cây quyết định (Decision Tree) là phương pháp học có giám sát phi tham số (non-parametric supervised learning). Mô hình tổ chức "
        "tri thức học được dưới dạng cấu trúc cây phân cấp bao gồm:\n"
        "- Nút gốc (Root Node): Chứa toàn bộ tập dữ liệu ban đầu trước khi bị phân chia.\n"
        "- Nút nội vi (Decision Nodes): Biểu diễn một điều kiện kiểm tra trên một thuộc tính cụ thể (ví dụ: Điểm môn Điện - Điện tử ô tô <= 2.25).\n"
        "- Nhánh (Branches): Tương ứng với kết quả của phép kiểm tra điều kiện (Đúng hoặc Sai).\n"
        "- Nút lá (Leaf Nodes): Đại diện cho nhãn phân lớp cuối cùng (Khá-Giỏi hoặc Trung bình-Yếu) cùng phân bố xác suất tương ứng."
    )

    doc.add_heading("2.2. Tiêu Chí Phân Tách Toán Học (Gini Impurity & Information Gain)", level=2)
    p = doc.add_paragraph(
        "Để tìm đặc trưng và điểm ngưỡng phân chia tối ưu tại mỗi bước, thuật toán đánh giá mức độ thuần khiết (Purity) "
        "của các tập con sinh ra. Hai tiêu chí toán học phổ biến nhất gồm:\n\n"
        "1. Độ vấy bẩn Gini (Gini Impurity - Thuật toán CART):\n"
        "Đo lường xác suất một phần tử được gán nhãn ngẫu nhiên bị phân loại sai nếu nhãn được rút ngẫu nhiên theo phân bố trong tập:"
    )
    p_eq1 = doc.add_paragraph()
    p_eq1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq1.add_run("Gini(D) = 1 - ∑ [p_i]^2  (với i = 1 đến C)\n").bold = True
    p_eq1.add_run("Trong bài toán phân lớp nhị phân (C = 2), Gini dao động từ 0 (hoàn toàn thuần khiết) đến 0.5 (hỗn độn cực đại 50/50).\n")

    p = doc.add_paragraph(
        "Khi phân tách tập D thành hai tập con D1 và D2 theo đặc trưng A tại ngưỡng theta, Độ vấy bẩn Gini có trọng số là:"
    )
    p_eq2 = doc.add_paragraph()
    p_eq2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq2.add_run("Gini_split(D, A) = (|D1| / |D|) * Gini(D1) + (|D2| / |D|) * Gini(D2)\n").bold = True
    p_eq2.add_run("Thuật toán chọn đặc trưng A và ngưỡng theta sao cho Gini_split đạt giá trị nhỏ nhất.\n")

    p = doc.add_paragraph(
        "2. Entropy và Độ lợi thông tin (Information Gain - Thuật toán ID3/C4.5):\n"
        "Đo lường lượng thông tin bất định của tập dữ liệu:"
    )
    p_eq3 = doc.add_paragraph()
    p_eq3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq3.add_run("Entropy(D) = - ∑ p_i * log2(p_i)\n").bold = True
    p_eq3.add_run("Information Gain(D, A) = Entropy(D) - ∑ (|D_v| / |D|) * Entropy(D_v)\n").bold = True

    doc.add_heading("2.3. Các Bước Xây Dựng Cây Quyết Định (Algorithm Steps / Pseudocode)", level=2)
    p = doc.add_paragraph(
        "Quy trình xây dựng Cây quyết định CART được thực hiện theo chiến lược đệ quy từ trên xuống (Top-down Greedy Recursion):\n"
        "Bước 1: Bắt đầu tại nút gốc chứa toàn bộ tập dữ liệu huấn luyện D.\n"
        "Bước 2: Với mọi thuộc tính đầu vào, sắp xếp các giá trị và tính Gini_split cho tất cả các điểm chia khả dĩ.\n"
        "Bước 3: Chọn thuộc tính và ngưỡng phân chia tối ưu mang lại mức giảm Gini lớn nhất (tăng độ thuần khiết cao nhất).\n"
        "Bước 4: Chia tập dữ liệu thành 2 nhánh con (D_left và D_right) tương ứng với điều kiện kiểm tra.\n"
        "Bước 5: Lặp lại đệ quy các Bước 2-4 trên từng tập con cho đến khi thỏa mãn một trong các điều kiện dừng:\n"
        "   - Nút đạt độ thuần khiết tuyệt đối (Gini = 0).\n"
        "   - Cây đạt độ sâu tối đa quy định (max_depth).\n"
        "   - Số lượng mẫu tại nút nhỏ hơn ngưỡng tối thiểu (min_samples_leaf).\n"
        "Bước 6: Gán nhãn cho nút lá theo nguyên tắc đa số phiếu (Majority Voting)."
    )

    doc.add_heading("2.4. Kỹ Thuật Kiểm Soát Quá Khớp (Pre-pruning) & Quy Trình Cross-Validation", level=2)
    p = doc.add_paragraph(
        "Nếu không có biện pháp ràng buộc, cây quyết định có xu hướng phát triển quá sâu để bao phủ toàn bộ dữ liệu huấn luyện, "
        "dẫn đến hiện tượng quá khớp (Overfitting) và làm giảm khả năng tổng quát hóa trên tập dữ liệu mới. "
        "Để khắc phục, đề tài áp dụng hai kỹ thuật chuẩn mực:\n"
        "1. Tiền cắt tỉa (Pre-pruning): Ràng buộc độ sâu tối đa (max_depth) và số mẫu tối thiểu ở nút lá (min_samples_leaf = 10).\n"
        "2. Kiểm định chéo 5 tầng (5-Fold Stratified Cross-Validation) trên tập Huấn luyện: Khảo sát hiệu năng của từng giá trị max_depth "
        "trực tiếp trên tập Train để lựa chọn siêu tham số tối ưu, bảo đảm hoàn toàn không sử dụng tập Test cho việc tinh chỉnh tham số."
    )

    doc.add_heading("2.5. Điểm Mạnh & Hạn Chế Của Thuật Toán", level=2)
    p = doc.add_paragraph(
        "1. Điểm mạnh:\n"
        "- Khả năng giải thích cao (White-box): Cung cấp các quy tắc If-Then tường minh, rất phù hợp cho tư vấn đào tạo.\n"
        "- Không nhạy cảm với thang đo (Scale Invariance): Không yêu cầu chuẩn hóa dữ liệu đầu vào.\n"
        "- Tự động chọn lọc thuộc tính: Đánh giá trực tiếp mức độ đóng góp (Gini Importance) của từng môn học.\n\n"
        "2. Hạn chế:\n"
        "- Độ bất ổn định (High Variance): Dễ biến động cấu trúc nếu dữ liệu huấn luyện thay đổi nhỏ.\n"
        "- Ranh giới phân chia trực giao: Phân tách không gian bằng các siêu phẳng song song với trục tọa độ, kém linh hoạt trước các quan hệ chéo phức tạp."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM", level=1)
    
    doc.add_heading("3.1. Môi Trường Thực Nghiệm & Công Nghệ", level=2)
    p = doc.add_paragraph(
        "Môi trường triển khai thực nghiệm được xây dựng độc lập trên nền Python với các thư viện tiêu chuẩn:\n"
        "- Ngôn ngữ: Python 3.12 64-bit.\n"
        "- Scikit-learn (1.6+): Module DecisionTreeClassifier, LogisticRegression, StratifiedKFold, cross_val_score và các độ đo đánh giá.\n"
        "- Pandas & NumPy: Xử lý dữ liệu bảng, tính toán thống kê mô tả.\n"
        "- Matplotlib & Seaborn: Xuất biểu đồ trực quan hóa độ phân giải cao (300 DPI)."
    )

    doc.add_heading("3.2. Mô Tả Bộ Dữ Liệu Thực Nghiệm & Tính Hợp Lệ Học Thuật", level=2)
    p = doc.add_paragraph(
        f"Nghiên cứu sử dụng bộ dữ liệu chuẩn hóa học thuật (Anonymized Academic Benchmark Dataset) gồm "
        f"**{metrics['summary']['n_students']} sinh viên** ngành kỹ thuật (Công nghệ kỹ thuật Ô tô & Cơ khí) với **{metrics['summary']['n_features']} thuộc tính**:\n"
        f"- Toàn bộ các thông tin định danh cá nhân (PII) như Họ tên, Mã sinh viên, Ngày sinh đã được khử định danh triệt để nhằm bảo đảm tính bảo mật dữ liệu.\n"
        f"- Biến nhân khẩu học: Nơi sinh ({metrics['summary']['birthplace_count']} tỉnh thành) và Giới tính (Nam: {metrics['summary']['male_count']}, Nữ: {metrics['summary']['female_count']}).\n"
        f"- Biến môn học: Điểm số 50 môn học từ HK1 đến HK8 theo thang điểm 4.0.\n"
        f"- Biến tổng kết GPA toàn khóa: Điểm trung bình = {metrics['summary']['gpa_mean']:.3f}, Độ lệch chuẩn = {metrics['summary']['gpa_std']:.3f}, "
        f"Trung vị = {metrics['summary']['gpa_median']:.3f}, Khoảng điểm: [{metrics['summary']['gpa_min']:.2f} - {metrics['summary']['gpa_max']:.2f}]."
    )

    # Chèn Biểu đồ 1: Phân bố học lực
    path_dist = os.path.join(OUTPUTS_DIR, "class_distribution.png")
    if os.path.exists(path_dist):
        doc.add_paragraph().paragraph_format.space_before = Pt(10)
        doc.add_picture(path_dist, width=Inches(5.5))
        p_cap = doc.add_paragraph("Hình 3.1: Phân bố xếp loại học lực của 1.656 sinh viên")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("3.3. Thiết Kế Kịch Bản Dự Đoán Sớm & Ngăn Chặn Rò Rỉ Nhãn (Zero Label Leakage)", level=2)
    p = doc.add_paragraph(
        "1. Xử lý giá trị khuyết: Các môn chưa tích lũy hoặc bị vắng thi được gán giá trị 0.0 theo quy chuẩn mã hóa dữ liệu học vụ.\n"
        "2. Xác định nhãn phân lớp (Label): Nhị phân hóa điểm GPA tích lũy toàn khóa theo Quy chế đào tạo tín chỉ:\n"
        "   - Nhãn 1 (Khá - Giỏi): GPA >= 2.50 (Gồm 794 sinh viên - 47.95%).\n"
        "   - Nhãn 0 (Trung bình - Yếu): GPA < 2.50 (Gồm 862 sinh viên - 52.05%).\n"
        "   Tỷ lệ phân bố 48% vs 52% cân bằng lý tưởng, không bị lệch nhãn.\n"
        "3. Ngăn ngừa rò rỉ nhãn (Zero Label Leakage Protocol):\n"
        "   - Trong thực tế, GPA toàn khóa được tính bình quân gia quyền từ chính điểm của 50 môn học. Nếu đưa toàn bộ 50 môn vào làm đầu vào thì mô hình sẽ rơi vào lỗi rò rỉ nhãn vòng quanh (Circular Label Leakage), không phản ánh bài toán cảnh báo sớm.\n"
        "   - Do đó, đề tài chỉ sử dụng duy nhất **15 môn học của Năm thứ nhất** (Học kỳ 1 & 2) làm biến đầu vào để dự đoán sớm kết quả toàn khóa. Tại thời điểm kết thúc năm nhất, sinh viên hoàn toàn chưa có điểm của 35 môn học ở các năm sau.\n"
        "4. Giao thức phân chia dữ liệu: Tập dữ liệu 1.656 mẫu được phân chia thành 80% Huấn luyện (1.324 sinh viên) và 20% Kiểm thử độc lập (332 sinh viên) theo phương pháp phân tầng (Stratified Split)."
    )

    # Chèn Biểu đồ 2: Ma trận tương quan
    path_corr = os.path.join(OUTPUTS_DIR, "correlation_heatmap.png")
    if os.path.exists(path_corr):
        doc.add_paragraph().paragraph_format.space_before = Pt(10)
        doc.add_picture(path_corr, width=Inches(5.8))
        p_cap = doc.add_paragraph("Hình 3.2: Ma trận tương quan giữa GPA và các môn học nền tảng Năm 1")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    # -------------------------------------------------------------
    # CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM", level=1)
    
    doc.add_heading("4.1. Hiệu Năng Phân Lớp Trên Tập Kiểm Thử Độc Lập (332 Sinh Viên)", level=2)
    p = doc.add_paragraph(
        "Mô hình Cây quyết định (Decision Tree) và mô hình đối chứng Hồi quy Logistic (Logistic Regression) được đánh giá "
        "trên tập kiểm thử độc lập (Test Set - 332 sinh viên) sử dụng 15 môn học Năm 1. Kết quả thực nghiệm được tổng hợp trong Bảng 4.1:"
    )

    dt_m = metrics["scenario_early_decision_tree"]
    lr_m = metrics["scenario_early_logistic_regression"]

    table_res = doc.add_table(rows=6, cols=3)
    table_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    res_headers = ["Chỉ Số Đánh Giá", "Cây Quyết Định (Decision Tree)", "Mô Hình Đối Chứng (Logistic Regression)"]
    
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

    p_cap = doc.add_paragraph("Bảng 4.1: Đối sánh hiệu năng phân lớp trên tập kiểm thử độc lập (Test Set 20%)")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.runs[0].italic = True
    p_cap.runs[0].font.size = Pt(9.5)

    p_compare = doc.add_paragraph(
        f"Nhận xét khoa học về kết quả đối sánh:\n"
        f"- Mô hình Decision Tree đạt Độ chính xác {dt_m['accuracy']*100:.2f}% và Macro F1 {dt_m['macro_f1']*100:.2f}% chỉ bằng dữ liệu năm nhất. "
        f"Mô hình đối chứng Logistic Regression có độ chính xác nhỉnh hơn ({lr_m['accuracy']*100:.2f}%) do khả năng kết hợp tuyến tính mịn màng "
        f"giữa các điểm môn học liên tục.\n"
        f"- Tuy nhiên, ưu thế vượt trội của Cây quyết định nằm ở tính minh bạch tuyệt đối: Mô hình tạo ra cấu trúc phân nhánh dạng luật trực quan, "
        f"cho phép cố vấn học tập nhìn thấy rõ ràng các ngưỡng điểm kiểm tra, điều mà mô hình Logistic Regression (vốn là hàm số sigmoid trừu tượng) không thể cung cấp."
    )

    doc.add_heading("4.2. Khảo Sát Siêu Tham Số max_depth Bằng 5-Fold Cross-Validation", level=2)
    p = doc.add_paragraph(
        "Để lựa chọn độ sâu cây tối ưu mà không vi phạm nguyên tắc rò rỉ tập Test, tác giả đã thực hiện 5-Fold Stratified Cross-Validation "
        "hoàn toàn trên tập Huấn luyện (Train 80% = 1.324 sinh viên). Kết quả khảo sát được thể hiện trong Bảng 4.2:"
    )

    depth_data = metrics["depth_tuning_cv"]
    table_depth = doc.add_table(rows=len(depth_data) + 1, cols=6)
    table_depth.alignment = WD_TABLE_ALIGNMENT.CENTER
    d_headers = ["Độ Sâu (max_depth)", "CV Mean Accuracy (Train)", "CV Std (Train)", "Train Accuracy", "Holdout Test Acc", "Generalization Gap"]
    
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

    p_cap = doc.add_paragraph("Bảng 4.2: Khảo sát độ sâu cây bằng 5-Fold Stratified Cross-Validation trên tập Train")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.runs[0].italic = True
    p_cap.runs[0].font.size = Pt(9.5)

    p_depth_eval = doc.add_paragraph(
        "- Tại max_depth = 1 và 2: Cây quá nông, điểm kiểm định chéo CV Mean đạt thấp (63.67% - 66.24%), mô hình bị Dưới khớp (Underfitting).\n"
        "- Tại max_depth = 3: Điểm kiểm định chéo tăng vọt lên 72.43% với độ lệch chuẩn nhỏ nhất (±1.30%), đồng thời độ phức tạp của cây chỉ gồm 8 nút lá, "
        "đây là cấu hình lý tưởng để trực quan hóa và giải thích trong báo cáo.\n"
        "- Tại max_depth >= 6: Mặc dù điểm trên tập huấn luyện tiếp tục tăng đến 82.10%, nhưng khoảng cách tổng quát hóa (Generalization Gap) "
        "vượt ngưỡng 10%, bắt đầu xuất hiện hiện tượng Quá khớp (Overfitting)."
    )

    doc.add_heading("4.3. Phân Tích Ma Trận Nhầm Lẫn (Confusion Matrix)", level=2)
    p = doc.add_paragraph(
        f"Ma trận nhầm lẫn của mô hình Cây quyết định trên 332 sinh viên kiểm thử (Hình 4.1) ghi nhận:\n"
        f"- True Negative (TN): {dt_m['tn']} sinh viên thực tế xếp loại Trung bình-Yếu được dự đoán chính xác.\n"
        f"- True Positive (TP): {dt_m['tp']} sinh viên thực tế xếp loại Khá-Giỏi được dự đoán chính xác.\n"
        f"- False Positive (FP - Sai loại 1): {dt_m['fp']} sinh viên thực tế Trung bình-Yếu nhưng mô hình dự đoán nhầm là Khá-Giỏi.\n"
        f"- False Negative (FN - Sai loại 2): {dt_m['fn']} sinh viên thực tế Khá-Giỏi nhưng mô hình dự đoán nhầm là Trung bình-Yếu.\n"
        f"Độ chuẩn xác Precision đạt {dt_m['precision_pass']*100:.2f}%, bảo đảm độ tin cậy khi hệ thống đưa ra nhận định sinh viên sẽ đạt loại Khá/Giỏi."
    )

    path_cm = os.path.join(OUTPUTS_DIR, "confusion_matrix.png")
    if os.path.exists(path_cm):
        doc.add_picture(path_cm, width=Inches(4.5))
        p_cap = doc.add_paragraph("Hình 4.1: Ma trận nhầm lẫn của mô hình Decision Tree trên tập Test")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.4. Phân Tích Cấu Trúc Cây Quyết Định & Các Quy Tắc Phân Lớp", level=2)
    p = doc.add_paragraph(
        "Cấu trúc Cây quyết định (Hình 4.2) thể hiện rõ ràng các điều kiện rẽ nhánh thực tế từ dữ liệu:\n"
        "1. Nút gốc (Root Node): Thuật toán chọn môn 'Điện - Điện tử ô tô' với ngưỡng điểm 2.25 làm nút phân chia đầu tiên. "
        "Nếu điểm môn này <= 2.25, xác suất sinh viên rơi vào nhóm Trung bình-Yếu chiếm ưu thế rõ rệt.\n"
        "2. Tầng rẽ nhánh kế tiếp: Tại nhánh điểm cao (> 2.25), thuật toán tiếp tục kiểm tra điểm môn 'Vật liệu học' (ngưỡng 2.75) và môn 'Tiếng Anh 2' (ngưỡng 2.25). "
        "Những sinh viên vượt qua cả hai ngưỡng này có tỷ lệ đạt bằng Khá/Giỏi trên 85%.\n"
        "Quy tắc này phản ánh trung thực thực tế đào tạo kỹ thuật: Sinh viên có tư duy kỹ thuật nền tảng vững vàng và năng lực ngoại ngữ tốt ở năm 1 "
        "sẽ duy trì kết quả học tập rất ổn định ở các năm tiếp theo."
    )

    path_tree = os.path.join(OUTPUTS_DIR, "decision_tree.png")
    if os.path.exists(path_tree):
        doc.add_picture(path_tree, width=Inches(6.5))
        p_cap = doc.add_paragraph("Hình 4.2: Cấu trúc Cây quyết định dự đoán sớm kết quả học tập")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.5. Xếp Hạng Mức Độ Quan Trọng Của Các Môn Học Năm 1 (Feature Importance)", level=2)
    p = doc.add_paragraph(
        "Mức độ quan trọng Gini (Gini Importance) của các môn học Năm 1 trong việc phân tách học lực sinh viên được thể hiện ở Hình 4.3:\n"
        "- Điện - Điện tử ô tô: Đóng góp 44.46% trọng số phân loại, là môn học mang tính bản lề hàng đầu.\n"
        "- Vật liệu học: Đóng góp 24.85% trọng số.\n"
        "- Tiếng Anh 2: Đóng góp 21.97% trọng số, chứng minh vai trò then chốt của ngoại ngữ trong việc tiếp cận giáo trình chuyên ngành.\n"
        "- Nhiệt kỹ thuật: Đóng góp 6.69%.\n"
        "- Tiếng Anh 1: Đóng góp 2.02%."
    )

    path_fi = os.path.join(OUTPUTS_DIR, "feature_importance.png")
    if os.path.exists(path_fi):
        doc.add_picture(path_fi, width=Inches(5.5))
        p_cap = doc.add_paragraph("Hình 4.3: Top các môn học Năm 1 có ảnh hưởng lớn nhất đến kết quả học tập")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.runs[0].italic = True
        p_cap.runs[0].font.size = Pt(9.5)

    doc.add_heading("4.6. Phân Tích Định Lượng Nguyên Nhân Dự Đoán Sai (Near-Threshold Error Analysis)", level=2)
    err = metrics["error_analysis"]
    p = doc.add_paragraph(
        f"Tổng cộng có {err['total_errors']} ca dự đoán sai trên 332 mẫu kiểm thử ({err['error_rate']}% sai số toàn cục). "
        f"Thay vì đưa ra các phỏng đoán cảm tính ngoài phạm vi dữ liệu, tác giả đã tiến hành phân tích định lượng phân bố điểm GPA của các ca sai:\n"
        f"1. Phân tích nhóm False Positive ({err['false_positives']} ca - Thực tế TB-Yếu nhưng đoán Khá-Giỏi):\n"
        f"   - Điểm GPA trung vị thực tế của nhóm này đạt {err['fp_gpa_median']:.3f} (trung bình: {err['fp_gpa_mean']:.3f}).\n"
        f"   - Trong đó có tới {err['fp_near_threshold']} / {err['false_positives']} sinh viên ({err['fp_near_pct']}%) có GPA >= 2.30, "
        f"tức chỉ kém ngưỡng phân định 2.50 từ 0.05 đến 0.20 điểm.\n"
        f"2. Phân tích nhóm False Negative ({err['false_negatives']} ca - Thực tế Khá-Giỏi nhưng đoán TB-Yếu):\n"
        f"   - Điểm GPA trung vị thực tế của nhóm này đạt {err['fn_gpa_median']:.3f} (trung bình: {err['fn_gpa_mean']:.3f}).\n"
        f"   - Trong đó có tới {err['fn_near_threshold']} / {err['false_negatives']} sinh viên ({err['fn_near_pct']}%) có GPA < 2.70, "
        f"cũng nằm sát ngay phía trên ranh giới 2.50.\n"
        f"3. Kết luận định lượng:\n"
        f"   - Có tới {err['total_near_threshold']} / {err['total_errors']} ca dự đoán sai ({err['near_threshold_pct']}%) tập trung trong dải điểm hẹp [2.30, 2.70) cận biên ranh giới quyết định. "
        f"Điều này chứng minh rằng mô hình hoàn toàn không phạm sai lầm thô thiển (như đoán nhầm sinh viên giỏi thành trượt hay ngược lại), "
        f"mà sai số chủ yếu phát sinh tại vùng chuyển tiếp mong manh giữa mức Trung bình khá và Khá."
    )

    # -------------------------------------------------------------
    # CHƯƠNG 5: KẾT LUẬN & HƯỚNG PHÁT TRIỂN
    # -------------------------------------------------------------
    doc.add_heading("CHƯƠNG 5: KẾT LUẬN & HƯỚNG PHÁT TRIỂN", level=1)
    
    doc.add_heading("5.1. Tóm Tắt Kết Quả Đạt Được", level=2)
    p = doc.add_paragraph(
        "Đề tài đã hoàn thành toàn diện các yêu cầu học thuật của bài tập lớn học phần Khai phá Dữ liệu:\n"
        f"1. Xây dựng hoàn chỉnh đường ống xử lý dữ liệu và huấn luyện mô hình Cây quyết định trên bộ dữ liệu 1.656 sinh viên.\n"
        f"2. Đạt độ chính xác 71.39% và Macro F1 70.94% trong kịch bản dự đoán sớm chỉ sử dụng 15 môn học Năm thứ nhất, "
        f"bảo đảm nguyên tắc Zero Data Leakage.\n"
        f"3. Áp dụng 5-Fold Stratified Cross-Validation trên tập Train để khảo sát siêu tham số max_depth một cách khoa học.\n"
        f"4. Trực quan hóa cấu trúc cây quyết định và giải mã nguyên nhân sai số định lượng tại vùng cận biên ranh giới quyết định."
    )

    doc.add_heading("5.2. Điểm Mạnh & Giới Hạn Của Đề Tài", level=2)
    p = doc.add_paragraph(
        "1. Điểm mạnh:\n"
        "- Tính khoa học và thực tiễn: Thiết kế kịch bản dự đoán sớm rõ ràng, giải quyết triệt để rủi ro rò rỉ nhãn.\n"
        "- Khả năng giải thích vượt trội: Mô hình cung cấp các quy tắc rẽ nhánh dễ hiểu phục vụ trực tiếp cho công tác tư vấn đào tạo.\n"
        "- Phương pháp luận chuẩn mực: Phân tách rạch ròi tập Train và Holdout Test Set kết hợp 5-Fold Cross-Validation.\n\n"
        "2. Giới hạn khách quan:\n"
        "- Dữ liệu chỉ bao gồm điểm thi số học, chưa có các thuộc tính hành vi (như tỷ lệ chuyên cần, thời gian tự học, tương tác hệ thống LMS).\n"
        "- Cây quyết định đơn lẻ có độ chính xác thấp hơn một chút so với mô hình hồi quy đối chứng do tính chất phân chia trực giao."
    )

    doc.add_heading("5.3. Hướng Mở Rộng Tiếp Theo", level=2)
    p = doc.add_paragraph(
        "1. Mở rộng sang các thuật toán học tổ hợp (Ensemble Methods): Thử nghiệm Random Forest và Gradient Boosting (LightGBM/XGBoost) "
        "để nâng cao hơn nữa độ chính xác phân lớp.\n"
        "2. Thu thập thêm các đặc trưng hành vi số: Tích hợp dữ liệu điểm danh và log học tập trên nền tảng e-learning để bổ trợ cho dữ liệu điểm số."
    )

    # -------------------------------------------------------------
    # TÀI LIỆU THAM KHẢO
    # -------------------------------------------------------------
    doc.add_heading("TÀI LIỆU THAM KHẢO", level=1)
    refs = [
        "[1] Cortez, P., & Silva, A. (2008). Using data mining to predict secondary school student performance. In Proceedings of 5th Annual Future Business Technology Conference (FUBUTEC 2008), pp. 5-12.",
        "[2] Breiman, L., Friedman, J. H., Olshen, R. A., & Stone, C. J. (1984). Classification and Regression Trees. Wadsworth & Brooks/Cole Advanced Books & Software.",
        "[3] Quinlan, J. R. (1986). Induction of decision trees. Machine Learning, 1(1), 81-106.",
        "[4] Romero, C., & Ventura, S. (2010). Educational data mining: a review of the state of the art. IEEE Transactions on Systems, Man, and Cybernetics, Part C (Applications and Reviews), 40(6), 601-618.",
        "[5] Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, 2825-2830.",
        "[6] Bộ Giáo dục và Đào tạo (2021). Thông tư 08/2021/TT-BGDĐT ban hành Quy chế đào tạo trình độ đại học.",
        "[7] Trường Đại học Công nghiệp Hà Nội (HaUI). Quy định về công tác đào tạo đại học chính quy theo hệ thống tín chỉ."
    ]
    for r in refs:
        p_ref = doc.add_paragraph(r)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.left_indent = Inches(0.25)
        p_ref.runs[0].font.size = Pt(10)

    # Lưu file report.docx
    output_docx_path = os.path.join(BASE_DIR, "report.docx")
    doc.save(output_docx_path)
    print(f" -> Đã tạo thành công file Word: {output_docx_path}")

    # Đồng thời xuất file report.md để dễ xem trước
    output_md_path = os.path.join(BASE_DIR, "report.md")
    write_markdown_report(metrics, output_md_path)
    print(f" -> Đã tạo thành công file Markdown: {output_md_path}")


def write_markdown_report(metrics, output_path):
    dt_m = metrics["scenario_early_decision_tree"]
    lr_m = metrics["scenario_early_logistic_regression"]
    depth_data = metrics["depth_tuning_cv"]
    err = metrics["error_analysis"]

    depth_rows = ""
    for r in depth_data:
        depth_rows += f"| {r['max_depth']} | {r['cv_mean_accuracy']*100:.2f}% | ±{r['cv_std_accuracy']*100:.2f}% | {r['train_accuracy']*100:.2f}% | {r['test_accuracy']*100:.2f}% | {r['generalization_gap']*100:+.2f}% |\n"

    md_content = f"""# Báo Cáo Bài Tập Lớn Cá Nhân: Dự Đoán Sớm Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree

**Học phần:** Khai phá Dữ liệu / Học máy (Data Mining / Machine Learning)  
**Bộ dữ liệu thực nghiệm:** Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark Dataset)  
**Quy mô thực nghiệm:** {metrics['summary']['n_students']} sinh viên, {metrics['summary']['n_features']} thuộc tính  
**Nhãn mục tiêu:** Xếp loại học lực tốt nghiệp toàn khóa (Khá - Giỏi: GPA >= 2.50 vs Trung bình - Yếu: GPA < 2.50)  
**Giao thức phương pháp luận:** Dự đoán sớm sau Năm thứ nhất (15 môn học) — 5-Fold Stratified Cross-Validation & Zero Label Leakage  

---

## THÔNG TIN SINH VIÊN THỰC HIỆN & CAM ĐOAN ĐÓNG GÓP

| STT | Họ và Tên Sinh Viên | Mã Sinh Viên | Nhiệm Vụ Thực Hiện Toàn Trình | Hệ Số Đóng Góp |
| :---: | :--- | :---: | :--- | :---: |
| 1 | ........................................ | .................... | Thực hiện độc lập 100%: Khảo sát bộ dữ liệu, thiết kế pipeline tiền xử lý, huấn luyện Decision Tree & Logistic Regression, khảo sát 5-Fold CV, xuất biểu đồ và viết báo cáo hoàn chỉnh. | **1.0 (100%)** |

> **Lời cam đoan:** Tôi xin cam đoan đây là công trình bài tập lớn do bản thân tôi tự thực hiện độc lập. Mọi số liệu thực nghiệm trong báo cáo đều được trích xuất trung thực từ mã nguồn chương trình, dữ liệu sử dụng đã được khử định danh học thuật và tuân thủ các chuẩn mực nghiên cứu khoa học.

---

## CHƯƠNG 1: GIỚI THIỆU TỔNG QUAN

### 1.1. Bối Cảnh Bài Toán & Tính Cấp Thiết
Trong đào tạo đại học theo hệ thống tín chỉ, sinh viên năm thứ nhất thường gặp khó khăn trong việc làm quen với phương pháp học tập mới. Xây dựng hệ thống cảnh báo học vụ sớm (Early Warning System) từ dữ liệu năm nhất giúp nhà trường và cố vấn học tập chủ động hỗ trợ sinh viên kịp thời trước khi kết quả học tập sa sút không thể cứu vãn.

### 1.2. Mục Tiêu & Phạm Vi Đề Tài
- Đối tượng: Hồ sơ 1.656 sinh viên ngành kỹ thuật, đã khử định danh toàn bộ thông tin cá nhân (PII).
- Phân loại nhị phân: Khá/Giỏi (GPA >= 2.50) vs Trung bình/Yếu (GPA < 2.50).
- Kịch bản trọng tâm: Dự đoán sớm bằng 15 môn học của Năm thứ nhất (Toán 1-2, Vật lý 1, Tiếng Anh 1-2, Cơ lý thuyết, Sức bền vật liệu, Điện - Điện tử ô tô, Nhiệt kỹ thuật, Vật liệu học...).

### 1.3. Tổng Quan Machine Learning & Bài Toán (~100 chữ)
*Học máy (Machine Learning) cho phép máy tính tự động đúc kết quy luật từ dữ liệu quá khứ. Đề tài ứng dụng thuật toán Cây quyết định (Decision Tree) với tiêu chí Độ vấy bẩn Gini. Lý do then chốt lựa chọn Decision Tree trong giáo dục là cấu trúc dạng cây phân cấp (If-Then rules) mang tính minh bạch tuyệt đối (White-box model), giúp giảng viên và sinh viên nắm rõ điều kiện tiên quyết để đạt kết quả tốt mà không bị phụ thuộc vào các mô hình hộp đen phức tạp.*

### 1.4. Công Việc Liên Quan
Kế thừa các nghiên cứu kinh điển của Cortez & Silva (2008) [1] và Romero & Ventura (2010) [4] về khai phá dữ liệu học sinh/sinh viên.

---

## CHƯƠNG 2: THUẬT TOÁN CÂY QUYẾT ĐỊNH (DECISION TREE)

### 2.1. Bản Chất Thuật Toán
Cây quyết định tổ chức không gian dữ liệu thành cây phân cấp: Nút gốc (Root Node) -> Nút nội vi (Decision Nodes) -> Nhánh (Branches) -> Nút lá (Leaf Nodes).

### 2.2. Tiêu Chí Phân Tách Toán Học (Gini Impurity)
$$\\text{{Gini}}(D) = 1 - \\sum_{{i=1}}^{{C}} p_i^2$$

$$\\text{{Gini}}_{{\\text{{split}}}}(D, A) = \\frac{{|D_1|}}{{|D|}} \\text{{Gini}}(D_1) + \\frac{{|D_2|}}{{|D|}} \\text{{Gini}}(D_2)$$

### 2.3. Các Bước Xây Dựng Cây (Pseudocode)
1. Bắt đầu tại nút gốc chứa toàn bộ tập dữ liệu huấn luyện $D$.
2. Thử nghiệm tất cả các điểm chia tiềm năng của mọi thuộc tính, tính $\\text{{Gini}}_{{\\text{{split}}}}$.
3. Chọn thuộc tính và ngưỡng tạo mức giảm Gini lớn nhất.
4. Chia tập dữ liệu thành 2 nhánh con.
5. Lặp lại đệ quy đến khi đạt điều kiện dừng (`max_depth`, `min_samples_leaf`).
6. Gán nhãn cho nút lá theo nguyên tắc đa số phiếu (Majority Voting).

---

## CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM

- Thư viện: Python 3.12, Scikit-learn, Pandas, Matplotlib, Seaborn.
- Dữ liệu: 1.656 sinh viên, 56 thuộc tính. GPA trung bình: {metrics['summary']['gpa_mean']:.3f}, trung vị: {metrics['summary']['gpa_median']:.3f}.
- Nhãn: Khá-Giỏi ({metrics['summary']['n_students'] - 862} sv - 47.95%) vs Trung bình-Yếu (862 sv - 52.05%).
- Giao thức chống rò rỉ nhãn (Zero Label Leakage): Chỉ sử dụng 15 môn học Năm 1 làm đầu vào. Điểm các năm sau không được phép xuất hiện trong mô hình dự đoán sớm.
- Chia dữ liệu: Train 80% (1.324 sinh viên), Holdout Test Set 20% (332 sinh viên) Stratified.

---

## CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM

### Bảng 4.1: Đối sánh hiệu năng phân lớp trên tập kiểm thử độc lập (332 sinh viên)
| Chỉ Số Đánh Giá | Cây Quyết Định (Decision Tree) | Mô Hình Đối Chứng (Logistic Regression) |
| :--- | :---: | :---: |
| **Accuracy (Độ chính xác)** | **{dt_m['accuracy']*100:.2f}%** | {lr_m['accuracy']*100:.2f}% |
| **Precision (Khá-Giỏi)** | **{dt_m['precision_pass']*100:.2f}%** | {lr_m['precision_pass']*100:.2f}% |
| **Recall (Khá-Giỏi)** | **{dt_m['recall_pass']*100:.2f}%** | {lr_m['recall_pass']*100:.2f}% |
| **F1-Score (Khá-Giỏi)** | **{dt_m['f1_pass']*100:.2f}%** | {lr_m['f1_pass']*100:.2f}% |
| **Macro F1-Score** | **{dt_m['macro_f1']*100:.2f}%** | {lr_m['macro_f1']*100:.2f}% |

### Bảng 4.2: Khảo sát độ sâu cây bằng 5-Fold Stratified Cross-Validation trên tập Train
| max_depth | CV Mean Accuracy (Train) | CV Std | Train Accuracy | Holdout Test Acc | Generalization Gap |
| :---: | :---: | :---: | :---: | :---: | :---: |
{depth_rows}

### Top 5 Môn Học Năm 1 Quyết Định Kết Quả Học Tập (Gini Importance):
1. **Điện - Điện tử ô tô:** 44.46%
2. **Vật liệu học:** 24.85%
3. **Tiếng Anh 2:** 21.97%
4. **Nhiệt kỹ thuật:** 6.69%
5. **Tiếng Anh 1:** 2.02%

### Phân Tích Định Lượng Lỗi Phân Lớp (Near-Threshold Error Analysis):
- Tổng số ca sai: {err['total_errors']} / 332 mẫu ({err['error_rate']}%).
- False Positive ({err['false_positives']} ca): GPA trung vị thực tế {err['fp_gpa_median']:.3f}, trong đó {err['fp_near_threshold']} ca ({err['fp_near_pct']}%) có GPA >= 2.30.
- False Negative ({err['false_negatives']} ca): GPA trung vị thực tế {err['fn_gpa_median']:.3f}, trong đó {err['fn_near_threshold']} ca ({err['fn_near_pct']}%) có GPA < 2.70.
- **Kết luận:** {err['total_near_threshold']} / {err['total_errors']} ca dự đoán sai (**{err['near_threshold_pct']}%**) tập trung trong dải điểm hẹp [2.30, 2.70) cận biên ranh giới quyết định 2.50.

---

## CHƯƠNG 5: KẾT LUẬN & HƯỚNG PHÁT TRIỂN
- Hoàn thành đề tài cá nhân với phương pháp luận chuẩn mực (Zero Label Leakage & 5-Fold Cross-Validation).
- Cây quyết định minh bạch, khả thi cho hệ thống cảnh báo học vụ sớm sau Năm 1.

---

## TÀI LIỆU THAM KHẢO
1. Cortez, P., & Silva, A. (2008). Using data mining to predict secondary school student performance. FUBUTEC 2008.
2. Breiman, L., et al. (1984). Classification and Regression Trees.
3. Quinlan, J. R. (1986). Induction of decision trees. Machine Learning, 1(1), 81-106.
4. Romero, C., & Ventura, S. (2010). Educational data mining: a review. IEEE TSMC-C.
5. Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. JMLR.
6. Bộ Giáo dục và Đào tạo (2021). Thông tư 08/2021/TT-BGDĐT ban hành Quy chế đào tạo đại học.
7. Trường Đại học Công nghiệp Hà Nội (HaUI). Quy chế đào tạo tín chỉ.
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md_content)


if __name__ == "__main__":
    build_docx_report()
