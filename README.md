# KPDL

## Dự Án Bài Tập Lớn Môn Khai Phá Dữ Liệu (KPDL)

### Đề tài: Dự Đoán Kết Quả Học Tập Của Sinh Viên Bằng Thuật Toán Decision Tree

- **Đơn vị thực hiện:** Sinh viên Trường Đại học Công nghiệp Hà Nội (HaUI)
- **Học phần:** Khai phá Dữ liệu (Data Mining) / Học máy (Machine Learning)
- **Bộ dữ liệu:** Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI)
- **Quy mô:** 1.656 sinh viên, 56 thuộc tính, 50 môn học từ HK1 đến HK8
- **Thuật toán chính:** Decision Tree Classifier (CART - Gini Impurity)
- **Mô hình đối chứng:** Logistic Regression

---

### Cấu Trúc Dự Án

Toàn bộ mã nguồn, dữ liệu thực nghiệm, các biểu đồ phân tích và file báo cáo Word/Markdown được đặt tại thư mục [`Student-Performance-Decision-Tree/`](./Student-Performance-Decision-Tree):

- [`Student-Performance-Decision-Tree/data/`](./Student-Performance-Decision-Tree/data): Chứa các tệp dữ liệu Excel gốc của sinh viên HaUI.
- [`Student-Performance-Decision-Tree/outputs/`](./Student-Performance-Decision-Tree/outputs): Chứa 5 biểu đồ độ phân giải cao (phân bố học lực, ma trận tương quan, hình cây quyết định, ma trận nhầm lẫn, mức độ quan trọng thuộc tính) và tệp `metrics.json`.
- [`Student-Performance-Decision-Tree/src/`](./Student-Performance-Decision-Tree/src): Các module xử lý dữ liệu (`data_loader.py`, `preprocessing.py`, `train_model.py`, `evaluate_model.py`).
- [`Student-Performance-Decision-Tree/main.py`](./Student-Performance-Decision-Tree/main.py): Kịch bản chạy pipeline tự động toàn trình.
- [`Student-Performance-Decision-Tree/generate_report_docx.py`](./Student-Performance-Decision-Tree/generate_report_docx.py): Script tự động xuất file Báo cáo Word.
- [`Student-Performance-Decision-Tree/report.docx`](./Student-Performance-Decision-Tree/report.docx): Báo cáo Bài tập lớn hoàn chỉnh định dạng Microsoft Word.
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
