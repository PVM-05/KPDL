# Báo Cáo Bài Tập Lớn Cá Nhân: Dự Đoán Sớm Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree

**Học phần:** Khai phá Dữ liệu / Học máy (Data Mining / Machine Learning)  
**Bộ dữ liệu thực nghiệm:** Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark Dataset)  
**Quy mô thực nghiệm:** 1656 sinh viên, 56 thuộc tính  
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
$$\text{Gini}(D) = 1 - \sum_{i=1}^{C} p_i^2$$

$$\text{Gini}_{\text{split}}(D, A) = \frac{|D_1|}{|D|} \text{Gini}(D_1) + \frac{|D_2|}{|D|} \text{Gini}(D_2)$$

### 2.3. Các Bước Xây Dựng Cây (Pseudocode)
1. Bắt đầu tại nút gốc chứa toàn bộ tập dữ liệu huấn luyện $D$.
2. Thử nghiệm tất cả các điểm chia tiềm năng của mọi thuộc tính, tính $\text{Gini}_{\text{split}}$.
3. Chọn thuộc tính và ngưỡng tạo mức giảm Gini lớn nhất.
4. Chia tập dữ liệu thành 2 nhánh con.
5. Lặp lại đệ quy đến khi đạt điều kiện dừng (`max_depth`, `min_samples_leaf`).
6. Gán nhãn cho nút lá theo nguyên tắc đa số phiếu (Majority Voting).

---

## CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM

- Thư viện: Python 3.12, Scikit-learn, Pandas, Matplotlib, Seaborn.
- Dữ liệu: 1.656 sinh viên, 56 thuộc tính. GPA trung bình: 2.399, trung vị: 2.485.
- Nhãn: Khá-Giỏi (794 sv - 47.95%) vs Trung bình-Yếu (862 sv - 52.05%).
- Giao thức chống rò rỉ nhãn (Zero Label Leakage): Chỉ sử dụng 15 môn học Năm 1 làm đầu vào. Điểm các năm sau không được phép xuất hiện trong mô hình dự đoán sớm.
- Chia dữ liệu: Train 80% (1.324 sinh viên), Holdout Test Set 20% (332 sinh viên) Stratified.

---

## CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM

### Bảng 4.1: Đối sánh hiệu năng phân lớp trên tập kiểm thử độc lập (332 sinh viên)
| Chỉ Số Đánh Giá | Cây Quyết Định (Decision Tree) | Mô Hình Đối Chứng (Logistic Regression) |
| :--- | :---: | :---: |
| **Accuracy (Độ chính xác)** | **71.39%** | 78.31% |
| **Precision (Khá-Giỏi)** | **74.24%** | 80.85% |
| **Recall (Khá-Giỏi)** | **61.64%** | 71.70% |
| **F1-Score (Khá-Giỏi)** | **67.35%** | 76.00% |
| **Macro F1-Score** | **70.94%** | 78.11% |

### Bảng 4.2: Khảo sát độ sâu cây bằng 5-Fold Stratified Cross-Validation trên tập Train
| max_depth | CV Mean Accuracy (Train) | CV Std | Train Accuracy | Holdout Test Acc | Generalization Gap |
| :---: | :---: | :---: | :---: | :---: | :---: |
| 1 | 63.67% | ±4.10% | 66.69% | 68.67% | -1.98% |
| 2 | 66.24% | ±1.69% | 68.58% | 68.67% | -0.09% |
| 3 | 72.43% | ±1.30% | 74.92% | 71.39% | +3.54% |
| 4 | 73.64% | ±2.21% | 76.13% | 74.40% | +1.74% |
| 5 | 71.60% | ±2.00% | 77.79% | 71.99% | +5.81% |
| 6 | 73.34% | ±2.29% | 80.21% | 71.39% | +8.83% |
| 7 | 74.09% | ±2.12% | 81.95% | 71.69% | +10.26% |
| 8 | 73.79% | ±2.82% | 82.10% | 71.39% | +10.71% |


### Top 5 Môn Học Năm 1 Quyết Định Kết Quả Học Tập (Gini Importance):
1. **Điện - Điện tử ô tô:** 44.46%
2. **Vật liệu học:** 24.85%
3. **Tiếng Anh 2:** 21.97%
4. **Nhiệt kỹ thuật:** 6.69%
5. **Tiếng Anh 1:** 2.02%

### Phân Tích Định Lượng Lỗi Phân Lớp (Near-Threshold Error Analysis):
- Tổng số ca sai: 95 / 332 mẫu (28.61%).
- False Positive (34 ca): GPA trung vị thực tế 2.388, trong đó 24 ca (70.59%) có GPA >= 2.30.
- False Negative (61 ca): GPA trung vị thực tế 2.595, trong đó 46 ca (75.41%) có GPA < 2.70.
- **Kết luận:** 70 / 95 ca dự đoán sai (**73.68%**) tập trung trong dải điểm hẹp [2.30, 2.70) cận biên ranh giới quyết định 2.50.

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
