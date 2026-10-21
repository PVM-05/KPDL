"""
Main Script: Dự Đoán Kết Quả Học Tập Sinh Viên Bằng Decision Tree
Nguồn dữ liệu: Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI)
Quy mô: 1.656 sinh viên, 56 biến (50 môn học trải dài từ HK1 đến HK8).

Quy trình 5 giai đoạn:
1. Đọc và khám phá dữ liệu (EDA)
2. Tiền xử lý dữ liệu & Chuẩn bị đặc trưng
3. Huấn luyện mô hình Cây quyết định & Mô hình đối chứng
4. Đánh giá thực nghiệm & Xuất 5 biểu đồ báo cáo
5. Lưu trữ số liệu thực nghiệm phục vụ xuất báo cáo hoàn chỉnh
"""

import os
import sys
import json
import pandas as pd
import numpy as np
from sklearn.metrics import confusion_matrix

# Thiết lập UTF-8 cho Windows console
if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Thiết lập thư mục gốc
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from src.data_loader import load_student_data, get_dataset_summary
from src.preprocessing import clean_course_grades, prepare_features, split_data
from src.train_model import train_decision_tree, train_baseline_model, tune_tree_depth
from src.evaluate_model import (
    calculate_metrics,
    plot_class_distribution,
    plot_correlation_heatmap,
    plot_decision_tree_graph,
    plot_confusion_matrix_heatmap,
    plot_feature_importance_bar,
    find_misclassifications
)


def main():
    print("=" * 75)
    print(" DỰ ĐOÁN KẾT QUẢ HỌC TẬP SINH VIÊN BẰNG THUẬT TOÁN DECISION TREE")
    print(" Nguồn dữ liệu: Đại học Công nghiệp Hà Nội (HaUI) | 1.656 sinh viên")
    print("=" * 75)

    data_dir = os.path.join(BASE_DIR, "data")
    outputs_dir = os.path.join(BASE_DIR, "outputs")
    os.makedirs(outputs_dir, exist_ok=True)

    # -------------------------------------------------------------
    # GIAI ĐOẠN 1: ĐỌC VÀ KHÁM PHÁ DỮ LIỆU
    # -------------------------------------------------------------
    print("\n[Giai đoạn 1] Đang nạp dữ liệu hồ sơ sinh viên HaUI...")
    df, df_dict = load_student_data(data_dir)
    summary = get_dataset_summary(df)
    
    print(f" -> Quy mô dữ liệu: {summary['n_students']} sinh viên | {summary['n_features']} biến")
    print(f" -> Điểm GPA tích lũy toàn khóa (Thang 4.0):")
    print(f"    * Trung bình (Mean):   {summary['gpa_mean']:.3f}")
    print(f"    * Độ lệch chuẩn (Std): {summary['gpa_std']:.3f}")
    print(f"    * Trung vị (Median):   {summary['gpa_median']:.3f}")
    print(f"    * Khoảng điểm:         {summary['gpa_min']:.2f} - {summary['gpa_max']:.2f}")

    # Xuất Biểu đồ 1 & 2
    path_dist = os.path.join(outputs_dir, "class_distribution.png")
    path_corr = os.path.join(outputs_dir, "correlation_heatmap.png")
    plot_class_distribution(df, path_dist)
    plot_correlation_heatmap(df, path_corr)
    print(f" -> Đã xuất biểu đồ phân bố học lực: {path_dist}")
    print(f" -> Đã xuất biểu đồ ma trận tương quan: {path_corr}")

    # -------------------------------------------------------------
    # GIAI ĐOẠN 2 & 3: TIỀN XỬ LÝ & HUẤN LUYỆN
    # Kịch bản 1: Dự đoán sớm (Chỉ dùng 15 môn học của Năm 1)
    # -------------------------------------------------------------
    print("\n[Giai đoạn 2 & 3] Huấn luyện Decision Tree (Kịch bản 1: Dự đoán sớm sau Năm thứ nhất)...")
    X_early, y = prepare_features(df, df_dict, scenario="early")
    X_train_e, X_test_e, y_train_e, y_test_e = split_data(X_early, y, test_size=0.2, random_state=42)

    print(f" -> Số đặc trưng Năm 1 sử dụng: {X_early.shape[1]} môn học nền tảng")
    print(f" -> Tập Huấn luyện (Train 80%): {X_train_e.shape[0]} sinh viên")
    print(f" -> Tập Kiểm thử (Test 20%):    {X_test_e.shape[0]} sinh viên")

    # Huấn luyện Decision Tree
    model_dt = train_decision_tree(X_train_e, y_train_e, max_depth=3, criterion="gini", min_samples_leaf=10, random_state=42)
    # Huấn luyện mô hình đối chứng Logistic Regression
    model_lr = train_baseline_model(X_train_e, y_train_e, random_state=42)

    # -------------------------------------------------------------
    # GIAI ĐOẠN 4: ĐÁNH GIÁ KẾT QUẢ THỰC NGHIỆM
    # -------------------------------------------------------------
    print("\n[Giai đoạn 4] Đánh giá trên tập kiểm thử độc lập (332 sinh viên)...")
    y_pred_dt = model_dt.predict(X_test_e)
    y_pred_lr = model_lr.predict(X_test_e)

    metrics_dt = calculate_metrics(y_test_e, y_pred_dt)
    metrics_lr = calculate_metrics(y_test_e, y_pred_lr)
    cm_dt = confusion_matrix(y_test_e, y_pred_dt)

    print("\n" + "=" * 55)
    print(" BẢNG SO SÁNH HIỆU QUẢ PHÂN LỚP (TẬP KIỂM THỬ 20%)")
    print("=" * 55)
    print(f"{'Chỉ Số Đánh Giá':<25} | {'Decision Tree':<12} | {'Logistic Reg':<12}")
    print("-" * 55)
    print(f"{'Accuracy (Độ chính xác)':<25} | {metrics_dt['accuracy']*100:>10.2f}% | {metrics_lr['accuracy']*100:>10.2f}%")
    print(f"{'Precision (Khá-Giỏi)':<25}   | {metrics_dt['precision_pass']*100:>10.2f}% | {metrics_lr['precision_pass']*100:>10.2f}%")
    print(f"{'Recall (Khá-Giỏi)':<25}      | {metrics_dt['recall_pass']*100:>10.2f}% | {metrics_lr['recall_pass']*100:>10.2f}%")
    print(f"{'F1-Score (Khá-Giỏi)':<25}    | {metrics_dt['f1_pass']*100:>10.2f}% | {metrics_lr['f1_pass']*100:>10.2f}%")
    print(f"{'Macro F1-Score':<25}        | {metrics_dt['macro_f1']*100:>10.2f}% | {metrics_lr['macro_f1']*100:>10.2f}%")
    print("=" * 55)

    # Xuất Biểu đồ 3, 4, 5
    path_tree = os.path.join(outputs_dir, "decision_tree.png")
    path_cm = os.path.join(outputs_dir, "confusion_matrix.png")
    path_fi = os.path.join(outputs_dir, "feature_importance.png")
    
    plot_decision_tree_graph(model_dt, X_early.columns.tolist(), path_tree)
    plot_confusion_matrix_heatmap(cm_dt, path_cm)
    df_fi = plot_feature_importance_bar(model_dt, X_early.columns.tolist(), path_fi, top_n=10)
    print(f" -> Đã xuất cấu trúc cây quyết định: {path_tree}")
    print(f" -> Đã xuất ma trận nhầm lẫn: {path_cm}")
    print(f" -> Đã xuất bảng xếp hạng môn học quan trọng: {path_fi}")

    # Khảo sát độ sâu cây (Tree Depth Tuning)
    depth_df = tune_tree_depth(X_train_e, y_train_e, X_test_e, y_test_e, range(1, 9))
    print("\n -> Khảo sát ảnh hưởng của max_depth (Tránh Overfitting):")
    print(depth_df.to_string(index=False))

    # Phân tích ca dự đoán sai (Error Analysis)
    errors_df = find_misclassifications(model_dt, X_test_e, y_test_e, df)
    print(f"\n -> Phân tích lỗi phân lớp:")
    print(f"    * Tổng số ca sai: {len(errors_df)} / {len(X_test_e)} ({len(errors_df)/len(X_test_e)*100:.2f}%)")
    fp_count = (errors_df["Actual_Label"] == 0) & (errors_df["Predicted_Label"] == 1)
    fn_count = (errors_df["Actual_Label"] == 1) & (errors_df["Predicted_Label"] == 0)
    print(f"    * False Positive (Thực tế TB-Yếu nhưng đoán Khá-Giỏi): {fp_count.sum()} sinh viên")
    print(f"    * False Negative (Thực tế Khá-Giỏi nhưng đoán TB-Yếu): {fn_count.sum()} sinh viên")

    # -------------------------------------------------------------
    # KỊCH BẢN ĐỐI CHỨNG: DỰ ĐOÁN TOÀN KHÓA (FULL PROFILE)
    # -------------------------------------------------------------
    print("\n[Đối chứng mở rộng] Đánh giá Kịch bản 2: Toàn bộ 50 môn học + Học bổng...")
    X_full, _ = prepare_features(df, df_dict, scenario="full")
    X_train_f, X_test_f, y_train_f, y_test_f = split_data(X_full, y, test_size=0.2, random_state=42)
    model_dt_full = train_decision_tree(X_train_f, y_train_f, max_depth=3, criterion="gini", min_samples_leaf=10, random_state=42)
    y_pred_f = model_dt_full.predict(X_test_f)
    metrics_full = calculate_metrics(y_test_f, y_pred_f)

    print(f" -> Accuracy (Toàn khóa): {metrics_full['accuracy']*100:.2f}% | F1-Score: {metrics_full['f1_pass']*100:.2f}%")

    # -------------------------------------------------------------
    # GIAI ĐOẠN 5: LƯU TRỮ METRICS CHO BÁO CÁO
    # -------------------------------------------------------------
    metrics_payload = {
        "dataset_name": "Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI)",
        "summary": summary,
        "scenario_early_decision_tree": metrics_dt,
        "scenario_early_logistic_regression": metrics_lr,
        "scenario_full_decision_tree": metrics_full,
        "depth_tuning": depth_df.to_dict(orient="records"),
        "top_features": df_fi.head(8).to_dict(orient="records"),
        "error_analysis": {
            "total_errors": int(len(errors_df)),
            "error_rate": round(float(len(errors_df)/len(X_test_e) * 100), 2),
            "false_positives": int(fp_count.sum()),
            "false_negatives": int(fn_count.sum())
        }
    }

    metrics_file = os.path.join(outputs_dir, "metrics.json")
    with open(metrics_file, "w", encoding="utf-8") as f:
        json.dump(metrics_payload, f, ensure_ascii=False, indent=4)

    print(f"\n[Hoàn thành] Đã lưu trữ toàn bộ chỉ số thực nghiệm vào: {metrics_file}")
    print("=" * 75)


if __name__ == "__main__":
    main()
