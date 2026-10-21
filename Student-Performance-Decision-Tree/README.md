# Dự Đoán Sớm Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree

Dự án ứng dụng kỹ thuật Khai phá Dữ liệu Giáo dục (Educational Data Mining - EDM) và thuật toán Cây quyết định (Decision Tree Classifier) nhằm dự đoán sớm xếp loại học lực của sinh viên đại học từ kết quả Năm thứ nhất, hỗ trợ công tác cố vấn học tập và cảnh báo học vụ.

---

## 1. Thông Tin Đề Tài & Bộ Dữ Liệu

- **Tên đề tài:** Dự đoán sớm kết quả học tập của sinh viên đại học bằng thuật toán Decision Tree (Cây quyết định).
- **Hình thức thực hiện:** Bài tập lớn cá nhân.
- **Môn học:** Khai phá Dữ liệu / Học máy (Data Mining / Machine Learning).
- **Bộ dữ liệu thực nghiệm:** **Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark Dataset)**.
  - Toàn bộ các thông tin định danh cá nhân (PII) đã được khử định danh triệt để nhằm bảo đảm tính riêng tư theo chuẩn nghiên cứu khoa học.
- **Quy mô dữ liệu:** **1.656 sinh viên**, **56 thuộc tính** bao gồm thông tin nhân khẩu học và điểm số chi tiết của các môn học theo thang điểm 4.0.
- **Biến mục tiêu (Label):** Xếp loại học lực tích lũy toàn khóa trên thang điểm 4.0 theo Quy chế đào tạo tín chỉ:
  - **Nhãn 1 (Khá - Giỏi):** $\text{GPA} \ge 2.50$ (794 sinh viên - 47.95%).
  - **Nhãn 0 (Trung bình - Yếu):** $\text{GPA} < 2.50$ (862 sinh viên - 52.05%).
- **Giao thức phương pháp luận (Zero Label Leakage):**
  - Chỉ sử dụng **15 môn học của Năm thứ nhất** (Học kỳ 1 & 2) làm biến đầu vào để dự đoán sớm.
  - Điểm các năm sau không được phép xuất hiện trong mô hình nhằm loại trừ hoàn toàn lỗi rò rỉ nhãn vòng quanh (Circular Label Leakage).
  - Áp dụng **5-Fold Stratified Cross-Validation trên tập Huấn luyện (Train 80%)** để khảo sát và chọn siêu tham số `max_depth`, bảo đảm tập Kiểm thử (Test 20%) độc lập hoàn toàn.

---

## 2. Cấu Trúc Thư Mục Dự Án

```text
Student-Performance-Decision-Tree/
│
├── data/
│   ├── 01_processed_dataset_labeled.xlsx       # Bộ dữ liệu 1.656 sinh viên với tên môn học tường minh
│   ├── 02_processed_dataset_harmonized_coded.xlsx # Bộ dữ liệu mã hóa chuẩn hóa
│   ├── 03_course_dictionary.xlsx              # Từ điển danh mục môn học
│   └── 04_variable_dictionary_and_summary.xlsx# Từ điển biến và thống kê mô tả
│
├── outputs/
│   ├── class_distribution.png                 # Biểu đồ 1: Phân bố xếp loại học lực sinh viên
│   ├── correlation_heatmap.png                # Biểu đồ 2: Ma trận tương quan giữa GPA và các môn Năm 1
│   ├── decision_tree.png                      # Biểu đồ 3: Cấu trúc Cây quyết định phân lớp
│   ├── confusion_matrix.png                   # Biểu đồ 4: Ma trận nhầm lẫn (Confusion Matrix)
│   ├── feature_importance.png                 # Biểu đồ 5: Xếp hạng mức độ quan trọng của các môn Năm 1
│   └── metrics.json                           # Toàn bộ số liệu thực nghiệm định lượng phục vụ báo cáo
│
├── src/
│   ├── __init__.py
│   ├── data_loader.py                         # Đọc và tóm tắt dữ liệu Excel
│   ├── preprocessing.py                       # Xử lý khuyết thiếu, gán nhãn, chia Train/Test 80/20 Stratified
│   ├── train_model.py                         # Huấn luyện Decision Tree, Logistic Regression, 5-Fold CV
│   └── evaluate_model.py                      # Tính toán metric, vẽ biểu đồ và phân tích lỗi cận biên
│
├── main.py                                    # Kịch bản thực thi tự động toàn trình 5 giai đoạn
├── generate_report_docx.py                    # Script tự động xuất file Báo cáo Word chuẩn mẫu giảng viên
├── report.docx                                # Báo cáo Bài tập lớn hoàn chỉnh (Microsoft Word)
├── report.md                                  # Báo cáo kỹ thuật chi tiết định dạng Markdown
├── requirements.txt                           # Danh sách các thư viện cần cài đặt
└── README.md                                  # Tài liệu hướng dẫn dự án
```

---

## 3. Hướng Dẫn Cài Đặt & Chạy Chương Trình

### Bước 1: Cài đặt môi trường Python
Yêu cầu Python phiên bản $\ge 3.9$. Cài đặt các thư viện phụ thuộc bằng lệnh:
```bash
pip install -r requirements.txt
```

### Bước 2: Chạy thực nghiệm tự động
Chỉ cần thực thi lệnh duy nhất:
```bash
python main.py
```
Toàn bộ quá trình nạp dữ liệu, tiền xử lý, kiểm định 5-Fold Cross-Validation, huấn luyện mô hình và xuất 5 biểu đồ phân giải cao sẽ hoàn thành trong khoảng 5–8 giây.

### Bước 3: Xuất file báo cáo Word (report.docx)
Chạy kịch bản xuất file báo cáo:
```bash
python generate_report_docx.py
```

---

## 4. Tóm Tắt Kết Quả Thực Nghiệm Thực Tế

### Bảng 1: So sánh hiệu năng phân lớp (Tập kiểm thử độc lập Test 20% = 332 sinh viên)
*Sử dụng 15 môn học nền tảng Năm thứ nhất*

| Chỉ số đánh giá | Thuật toán Decision Tree (Cây quyết định) | Thuật toán đối chứng (Logistic Regression) |
| :--- | :---: | :---: |
| **Độ chính xác (Accuracy)** | **71.39%** | 78.31% |
| **Precision (Lớp Khá - Giỏi)** | **74.24%** | 80.85% |
| **Recall (Lớp Khá - Giỏi)** | **61.64%** | 71.70% |
| **F1-Score (Lớp Khá - Giỏi)** | **67.35%** | 76.00% |
| **Macro F1-Score** | **70.94%** | 78.11% |

### Bảng 2: Khảo sát độ sâu cây bằng 5-Fold Stratified Cross-Validation trên tập Train
| Độ sâu (`max_depth`) | CV Mean Accuracy (Train) | CV Std | Train Accuracy | Holdout Test Acc | Generalization Gap |
| :---: | :---: | :---: | :---: | :---: | :---: |
| **1** | 63.67% | ±4.10% | 66.69% | 68.67% | -1.98% |
| **2** | 66.24% | ±1.69% | 68.58% | 68.67% | -0.09% |
| **3** | **72.43%** | **±1.30%** | **74.92%** | **71.39%** | **+3.54%** |
| **4** | 73.64% | ±2.21% | 76.13% | 74.40% | +1.74% |
| **5** | 71.60% | ±2.00% | 77.79% | 71.99% | +5.81% |
| **6** | 73.34% | ±2.29% | 80.21% | 71.39% | +8.83% |
| **7** | 74.09% | ±2.12% | 81.95% | 71.69% | +10.26% |
| **8** | 73.79% | ±2.82% | 82.10% | 71.39% | +10.71% |

### Bảng 3: Top môn học Năm 1 quyết định kết quả học tập (Gini Importance)
1. **Điện - Điện tử ô tô (Automotive Electrical and Electronics):** Chiếm **44.46%** trọng số phân nhánh.
2. **Vật liệu học (Materials Science):** Chiếm **24.85%** trọng số.
3. **Tiếng Anh 2 (English 2):** Chiếm **21.97%** trọng số.
4. **Nhiệt kỹ thuật (Engineering Thermodynamics):** Chiếm **6.69%** trọng số.
5. **Tiếng Anh 1 (English 1):** Chiếm **2.02%** trọng số.

### Bảng 4: Phân tích định lượng lỗi phân lớp (Near-Threshold Error Analysis)
- Tổng số ca sai: **95 / 332 sinh viên** (28.61%).
- False Positive (34 ca): Điểm GPA trung vị là **2.388**, trong đó có **24 ca (70.59%) có $\text{GPA} \ge 2.30$** (chỉ kém ngưỡng 2.50 từ 0.05 đến 0.20 điểm).
- False Negative (61 ca): Điểm GPA trung vị là **2.595**, trong đó có **46 ca (75.41%) có $\text{GPA} < 2.70$** (sát ngay trên ranh giới 2.50).
- **Kết luận:** Có tới **70 / 95 ca sai (73.68%)** tập trung ở dải điểm hẹp $[2.30, 2.70)$ cận biên ranh giới quyết định. Sai số chủ yếu phát sinh tại vùng chuyển tiếp mong manh giữa Trung bình khá và Khá, không có sai sót cực đoan.
