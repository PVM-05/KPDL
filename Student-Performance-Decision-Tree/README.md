# Dự Đoán Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree

Dự án ứng dụng kỹ thuật Khai phá Dữ liệu Giáo dục (Educational Data Mining - EDM) và thuật toán Cây quyết định (Decision Tree Classifier) nhằm dự đoán sớm xếp loại học lực của sinh viên đại học, hỗ trợ công tác cố vấn học tập và cảnh báo học vụ.

---

## 1. Thông Tin Đề Tài & Bộ Dữ Liệu

- **Tên đề tài:** Dự đoán kết quả học tập của sinh viên đại học bằng thuật toán Decision Tree (Cây quyết định).
- **Môn học:** Khai phá Dữ liệu / Học máy (Data Mining / Machine Learning).
- **Bộ dữ liệu thực nghiệm:** **Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI)**.
- **Quy mô dữ liệu:** **1.656 sinh viên**, **56 biến** bao gồm thông tin nhân khẩu học và điểm số chi tiết của 50 môn học trải dài từ Học kỳ 1 đến Học kỳ 8.
- **Biến mục tiêu (Label):** Xếp loại học lực tích lũy toàn khóa trên thang điểm 4.0 theo Quy chế đào tạo tín chỉ:
  - **Nhãn 1 (Khá - Giỏi):** $\text{GPA} \ge 2.50$ (794 sinh viên - 47.95%).
  - **Nhãn 0 (Trung bình - Yếu):** $\text{GPA} < 2.50$ (862 sinh viên - 52.05%).

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
│   ├── correlation_heatmap.png                # Biểu đồ 2: Ma trận tương quan giữa GPA và các môn nền tảng
│   ├── decision_tree.png                      # Biểu đồ 3: Cấu trúc Cây quyết định phân lớp
│   ├── confusion_matrix.png                   # Biểu đồ 4: Ma trận nhầm lẫn (Confusion Matrix)
│   ├── feature_importance.png                 # Biểu đồ 5: Xếp hạng mức độ quan trọng của các môn học
│   └── metrics.json                           # Toàn bộ số liệu thực nghiệm định lượng phục vụ báo cáo
│
├── src/
│   ├── __init__.py
│   ├── data_loader.py                         # Đọc và tóm tắt dữ liệu Excel
│   ├── preprocessing.py                       # Xử lý khuyết thiếu, gán nhãn, chia tập Train/Test (80/20)
│   ├── train_model.py                         # Huấn luyện Decision Tree, Logistic Regression, khảo sát max_depth
│   └── evaluate_model.py                      # Tính toán metric, vẽ biểu đồ và phân tích ca dự đoán sai
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
Toàn bộ quá trình nạp dữ liệu, tiền xử lý, huấn luyện 2 kịch bản, khảo sát độ sâu cây và xuất 5 biểu đồ độ phân giải cao sẽ hoàn thành trong khoảng 5–10 giây.

### Bước 3: Xuất file báo cáo Word (report.docx)
Chạy kịch bản xuất file báo cáo:
```bash
python generate_report_docx.py
```

---

## 4. Tóm Tắt Kết Quả Thực Nghiệm Thực Tế

### Bảng 1: So sánh hiệu năng phân lớp (Tập kiểm thử Test 20% = 332 sinh viên)
*Kịch bản 1: Dự đoán sớm sau Năm 1 (Chỉ sử dụng 15 môn học nền tảng)*

| Chỉ số đánh giá | Thuật toán Decision Tree (Cây quyết định) | Thuật toán đối chứng (Logistic Regression) |
| :--- | :---: | :---: |
| **Độ chính xác (Accuracy)** | **71.39%** | 78.31% |
| **Precision (Lớp Khá - Giỏi)** | **74.24%** | 80.85% |
| **Recall (Lớp Khá - Giỏi)** | **61.64%** | 71.70% |
| **F1-Score (Lớp Khá - Giỏi)** | **67.35%** | 76.00% |
| **Macro F1-Score** | **70.94%** | 78.11% |

### Bảng 2: Khảo sát ảnh hưởng của độ sâu cây (`max_depth`)
| Độ sâu (`max_depth`) | Train Accuracy | Test Accuracy | Khoảng cách Tổng quát hóa (Gap) | Đánh giá hiện tượng |
| :---: | :---: | :---: | :---: | :--- |
| **1** | 66.69% | 68.67% | -1.98% | Dưới khớp (Underfitting) |
| **2** | 68.58% | 68.67% | -0.09% | Cây quá nông, chưa bắt kịp quy luật |
| **3** | 74.92% | **71.39%** | +3.54% | **Điểm cân bằng, cây trong sáng dễ giải thích** |
| **4** | 76.13% | **74.40%** | +1.74% | **Độ chính xác Test cao nhất** |
| **5** | 77.79% | 71.99% | +5.81% | Bắt đầu phân nhánh dư thừa |
| **6** | 80.21% | 71.39% | +8.82% | Bắt đầu quá khớp (Overfitting) |
| **8** | 82.10% | 71.39% | +10.71% | Quá khớp rõ rệt trên tập huấn luyện |

### Bảng 3: Top môn học Năm 1 quyết định kết quả học tập (Gini Importance)
1. **Điện - Điện tử ô tô (Automotive Electrical and Electronics):** Chiếm **44.46%** trọng số phân nhánh.
2. **Vật liệu học (Materials Science):** Chiếm **24.85%** trọng số.
3. **Tiếng Anh 2 (English 2):** Chiếm **21.97%** trọng số.
4. **Nhiệt kỹ thuật (Engineering Thermodynamics):** Chiếm **6.69%** trọng số.
5. **Tiếng Anh 1 (English 1):** Chiếm **2.02%** trọng số.

*Nhận xét:* Sinh viên có kết quả vững chắc ở các môn chuyên ngành cơ sở (Điện tử ô tô, Vật liệu học) và ngoại ngữ (Tiếng Anh 1 & 2) ngay từ năm thứ nhất có tỷ lệ tốt nghiệp loại Khá/Giỏi cao vượt trội.
