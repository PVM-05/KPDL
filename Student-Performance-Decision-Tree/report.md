# Báo Cáo Bài Tập Lớn: Dự Đoán Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree

**Môn học:** Khai phá Dữ liệu / Học máy (Data Mining / Machine Learning)  
**Bộ dữ liệu thực nghiệm:** Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI)  
**Quy mô thực nghiệm:** 1656 sinh viên, 56 biến  
**Nhãn mục tiêu:** Xếp loại học lực tốt nghiệp toàn khóa (Khá - Giỏi: GPA >= 2.50 vs Trung bình - Yếu: GPA < 2.50)  

---

## BẢNG XÁC NHẬN MỨC ĐỘ ĐÓNG GÓP CỦA THÀNH VIÊN

| STT | Họ và Tên | Mã Sinh Viên | Nội Dung Phụ Trách & Nhiệm Vụ | Mức Đóng Góp |
| :---: | :--- | :---: | :--- | :---: |
| 1 | Thành viên 1 (Trưởng nhóm) | .................... | Khảo sát dữ liệu HaUI, viết mã nguồn tiền xử lý (clean, split), huấn luyện mô hình Decision Tree và viết Báo cáo Chương 1, 2, 3 | **1.0 (100%)** |
| 2 | Thành viên 2 | .................... | Xây dựng mô hình đối chứng Logistic Regression, khảo sát max_depth, phân tích ma trận nhầm lẫn, vẽ biểu đồ và viết Chương 4, 5 | **1.0 (100%)** |

> **Cam kết:** Chúng tôi đồng ý và xác nhận mức độ đóng góp cũng như kết quả phân chia nhiệm vụ như trên. Mọi số liệu thực nghiệm trong báo cáo đều được trích xuất trung thực từ mã nguồn chương trình.

---

## CHƯƠNG 1: GIỚI THIỆU TỔNG QUAN

### 1.1. Bối Cảnh Bài Toán & Tính Cấp Thiết
Trong giáo dục đại học theo hệ thống tín chỉ, việc sinh viên gặp bỡ ngỡ ở năm thứ nhất dẫn đến kết quả học tập sa sút, bị cảnh báo học vụ hoặc buộc thôi học là thực trạng phổ biến. Bài toán Khai phá Dữ liệu Giáo dục (Educational Data Mining - EDM) nhằm dự đoán sớm kết quả học tập từ dữ liệu năm nhất có ý nghĩa thực tiễn to lớn: giúp nhà trường và cố vấn học tập chủ động phát hiện sinh viên có nguy cơ để can thiệp kịp thời.

### 1.2. Mục Tiêu & Phạm Vi Đề Tài
- Đối tượng: 1.656 sinh viên ngành kỹ thuật thuộc Trường Đại học Công nghiệp Hà Nội (HaUI).
- Phân loại nhị phân: Khá/Giỏi (GPA >= 2.50) vs Trung bình/Yếu (GPA < 2.50).
- Kịch bản trọng tâm: Dự đoán sớm bằng 15 môn học của Năm thứ nhất.

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
2. Thử nghiệm tất cả các điểm chia tiềm năng của mọi đặc trưng, tính $\text{Gini}_{\text{split}}$.
3. Chọn đặc trưng và ngưỡng tạo mức giảm Gini lớn nhất.
4. Chia tập dữ liệu thành 2 nhánh con.
5. Lặp lại đệ quy đến khi đạt điều kiện dừng (`max_depth`, `min_samples_leaf`).
6. Gán nhãn cho nút lá theo nguyên tắc đa số phiếu (Majority Voting).

---

## CHƯƠNG 3: TRIỂN KHAI THỰC NGHIỆM

- Thư viện: Python 3.12, Scikit-learn, Pandas, Matplotlib, Seaborn.
- Dữ liệu HaUI: 1.656 sinh viên, 56 thuộc tính. GPA trung bình: 2.399, trung vị: 2.485.
- Nhãn: Khá-Giỏi (794 sv - 47.95%) vs Trung bình-Yếu (862 sv - 52.05%).
- Chia dữ liệu: Train 80% (1.324 sinh viên), Test 20% (332 sinh viên) Stratified.

---

## CHƯƠNG 4: PHÂN TÍCH KẾT QUẢ THỰC NGHIỆM

### Bảng 4.1: Đối sánh hiệu năng phân lớp (Tập kiểm thử 332 sinh viên)
| Chỉ Số Đánh Giá | Decision Tree (Năm 1) | Logistic Regression (Đối chứng) | Decision Tree (Toàn khóa) |
| :--- | :---: | :---: | :---: |
| **Accuracy (Độ chính xác)** | **71.39%** | 78.31% | 84.34% |
| **Precision (Khá-Giỏi)** | **74.24%** | 80.85% | 92.80% |
| **Recall (Khá-Giỏi)** | **61.64%** | 71.70% | 72.96% |
| **F1-Score (Khá-Giỏi)** | **67.35%** | 76.00% | 81.69% |
| **Macro F1-Score** | **70.94%** | 78.11% | 84.00% |

### Bảng 4.2: Khảo sát ảnh hưởng của max_depth (Tránh Overfitting)
| max_depth | Train Accuracy | Test Accuracy | Generalization Gap | Nhận xét |
| :---: | :---: | :---: | :---: | :--- |
| **1** | 66.69% | 68.67% | -1.98% | Dưới khớp (Underfitting) |
| **2** | 68.58% | 68.67% | -0.09% | Cây quá nông |
| **3** | 74.92% | **71.39%** | +3.54% | **Điểm cân bằng, cây trong sáng dễ giải thích** |
| **4** | 76.13% | **74.40%** | +1.74% | **Test Accuracy cao nhất** |
| **5** | 77.79% | 71.99% | +5.81% | Bắt đầu phân nhánh dư thừa |
| **6** | 80.21% | 71.39% | +8.82% | Bắt đầu quá khớp |
| **8** | 82.10% | 71.39% | +10.71% | Quá khớp rõ rệt (Overfitting) |

### Top 5 Môn Học Năm 1 Quyết Định Kết Quả Học Tập (Gini Importance):
1. **Điện - Điện tử ô tô:** 44.46%
2. **Vật liệu học:** 24.85%
3. **Tiếng Anh 2:** 21.97%
4. **Nhiệt kỹ thuật:** 6.69%
5. **Tiếng Anh 1:** 2.02%

---

## CHƯƠNG 5: KẾT LUẬN & HƯỚNG PHÁT TRIỂN
- Hoàn thành đầy đủ 5 mục tiêu đề ra với dữ liệu sinh viên thật của HaUI.
- Điểm mạnh: Cây quyết định minh bạch, khả thi cho bài toán cảnh báo học vụ sớm sau năm 1.
- Hướng mở rộng: Tích hợp Ensemble methods (Random Forest, LightGBM) và dữ liệu hệ thống học tập LMS.

---

## TÀI LIỆU THAM KHẢO
1. Cortez, P., & Silva, A. (2008). Using data mining to predict secondary school student performance. FUBUTEC 2008.
2. Breiman, L., et al. (1984). Classification and Regression Trees.
3. Quinlan, J. R. (1986). Induction of decision trees. Machine Learning, 1(1), 81-106.
4. Romero, C., & Ventura, S. (2010). Educational data mining: a review. IEEE TSMC-C.
5. Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. JMLR.
6. Bộ Giáo dục và Đào tạo (2021). Thông tư 08/2021/TT-BGDĐT ban hành Quy chế đào tạo đại học.
7. Trường Đại học Công nghiệp Hà Nội (HaUI). Quy chế đào tạo tín chỉ.
