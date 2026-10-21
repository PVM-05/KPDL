# KPDL

## Dự Án Bài Tập Lớn Cá Nhân Môn Khai Phá Dữ Liệu (KPDL)

### Đề tài: Dự Đoán Sớm Kết Quả Học Tập Của Sinh Viên Bằng Thuật Toán Decision Tree

- **Học phần:** Khai phá Dữ liệu (Data Mining) / Học máy (Machine Learning)
- **Hình thức thực hiện:** Bài tập lớn cá nhân (Hệ số đóng góp 1.0 - 100%)
- **Bộ dữ liệu:** Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark Dataset)
- **Quy mô:** 1.656 sinh viên, 56 thuộc tính
- **Phương pháp luận:** Kịch bản dự đoán sớm sau Năm thứ nhất (15 môn học nền tảng) — Kiểm soát 100% rò rỉ nhãn (Zero Label Leakage) & 5-Fold Stratified Cross-Validation trên tập Train
- **Thuật toán chính:** Decision Tree Classifier (CART - Gini Impurity)
- **Mô hình đối chứng:** Logistic Regression

---

### Cấu Trúc Dự Án

Toàn bộ mã nguồn, dữ liệu thực nghiệm, các biểu đồ phân tích và file báo cáo Word/Markdown được đặt tại thư mục [`Student-Performance-Decision-Tree/`](./Student-Performance-Decision-Tree):

- [`Student-Performance-Decision-Tree/data/`](./Student-Performance-Decision-Tree/data): Chứa các tệp dữ liệu Excel chuẩn hóa và từ điển biến.
- [`Student-Performance-Decision-Tree/outputs/`](./Student-Performance-Decision-Tree/outputs): Chứa 5 biểu đồ độ phân giải cao 300 DPI và tệp số liệu thực nghiệm [`metrics.json`](./Student-Performance-Decision-Tree/outputs/metrics.json).
- [`Student-Performance-Decision-Tree/src/`](./Student-Performance-Decision-Tree/src): Các module xử lý dữ liệu (`data_loader.py`, `preprocessing.py`, `train_model.py`, `evaluate_model.py`).
- [`Student-Performance-Decision-Tree/main.py`](./Student-Performance-Decision-Tree/main.py): Kịch bản chạy pipeline tự động toàn trình.
- [`Student-Performance-Decision-Tree/generate_report_docx.py`](./Student-Performance-Decision-Tree/generate_report_docx.py): Script tự động xuất file Báo cáo Word.
- [`Student-Performance-Decision-Tree/report.docx`](./Student-Performance-Decision-Tree/report.docx): Báo cáo Bài tập lớn hoàn chỉnh định dạng Microsoft Word (Đã nhúng bảng biểu & đồ thị).
- [`Student-Performance-Decision-Tree/report.md`](./Student-Performance-Decision-Tree/report.md): Báo cáo kỹ thuật chi tiết định dạng Markdown.

---

### Hướng Dẫn Chạy Chương Trình

1. Cài đặt các thư viện phụ thuộc:
```bash
pip install -r Student-Performance-Decision-Tree/requirements.txt
```

2. Chạy toàn bộ thực nghiệm và xuất biểu đồ:
```bash
python Student-Performance-Decision-Tree/main.py
```

3. Xuất file Báo cáo Word:
```bash
python Student-Performance-Decision-Tree/generate_report_docx.py
```
