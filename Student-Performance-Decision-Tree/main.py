"""
Main Script: Dự Đoán Kết Quả Học Tập Sinh Viên Bằng Thuật Toán Decision Tree
Tập trung 100% vào Kịch bản Dự đoán sớm sau Năm thứ nhất (Early Warning Prediction).
Quy trình thực nghiệm chuẩn khoa học: 5-Fold Stratified Cross-Validation & Zero Test Leakage.
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
from src.train_model import train_decision_tree, train_baseline_model, tune_tree_depth_cv
from src.evaluate_model import (
    calculate_metrics,
    plot_class_distribution,
    plot_correlation_heatmap,
    plot_decision_tree_graph,
    plot_confusion_matrix_heatmap,
    plot_feature_importance_bar,
    analyze_errors_quantitatively
)


def main():
    print("=" * 75)
    print(" DỰ ĐOÁN SỚM KẾT QUẢ HỌC TẬP SINH VIÊN BẰNG THUẬT TOÁN DECISION TREE")
    print(" Đối tượng: Hồ sơ 1.656 sinh viên | Zero Label Leakage Protocol")
    print("=" * 75)

    data_dir = os.path.join(BASE_DIR, "data")
    outputs_dir = os.path.join(BASE_DIR, "outputs")
    os.makedirs(outputs_dir, exist_ok=True)

    # -------------------------------------------------------------
    # GIAI ĐOẠN 1: ĐỌC VÀ KHÁM PHÁ DỮ LIỆU
    # -------------------------------------------------------------
    print("\n[Giai đoạn 1] Nạp dữ liệu hồ sơ học tập đã khử định danh...")
    df, df_dict = load_student_data(data_dir)
    summary = get_dataset_summary(df)
    
    print(f" -> Quy mô dữ liệu: {summary['n_students']} sinh viên | {summary['n_features']} thuộc tính")
    print(f" -> Thống kê GPA tích lũy toàn khóa (Thang 4.0):")
    print(f"    * Điểm trung bình (Mean):   {summary['gpa_mean']:.3f}")
    print(f"    * Độ lệch chuẩn (Std):     {summary['gpa_std']:.3f}")
    print(f"    * Trung vị (Median):       {summary['gpa_median']:.3f}")
    print(f"    * Khoảng điểm:             {summary['gpa_min']:.2f} - {summary['gpa_max']:.2f}")

    # Xuất Biểu đồ 1 & 2
    path_dist = os.path.join(outputs_dir, "class_distribution.png")
    path_corr = os.path.join(outputs_dir, "correlation_heatmap.png")
    plot_class_distribution(df, path_dist)
    plot_correlation_heatmap(df, path_corr)
    print(f" -> Đã xuất biểu đồ phân bố học lực: {path_dist}")
    print(f" -> Đã xuất biểu đồ ma trận tương quan: {path_corr}")

    # -------------------------------------------------------------
    # GIAI ĐOẠN 2: TIỀN XỬ LÝ & THIẾT KẾ KỊCH BẢN DỰ ĐOÁN SỚM
    # -------------------------------------------------------------
    print("\n[Giai đoạn 2] Chuẩn bị đặc trưng (Kịch bản Dự đoán sớm - 15 môn học Năm 1)...")
    X_early, y = prepare_features(df, df_dict, scenario="early")
    X_train_e, X_test_e, y_train_e, y_test_e = split_data(X_early, y, test_size=0.2, random_state=42)

    print(f" -> Số đặc trưng Năm 1 sử dụng: {X_early.shape[1]} môn học nền tảng")
    print(f" -> Tập Huấn luyện (Train 80%): {X_train_e.shape[0]} sinh viên")
    print(f" -> Tập Kiểm thử (Holdout Test 20%): {X_test_e.shape[0]} sinh viên (Độc lập 100%)")

    # -------------------------------------------------------------
    # GIAI ĐOẠN 3: KHẢO SÁT 5-FOLD CV TRÊN TẬP TRAIN & HUẤN LUYỆN
    # -------------------------------------------------------------
    print("\n[Giai đoạn 3] Khảo sát max_depth bằng 5-Fold Stratified Cross-Validation trên Train...")
    depth_df = tune_tree_depth_cv(X_train_e, y_train_e, X_test_e, y_test_e, range(1, 9), cv=5, random_state=42)
    print(depth_df.to_string(index=False))

    # Huấn luyện mô hình Cây quyết định chính (max_depth=3 để cây sáng sủa, dễ giải thích trong báo cáo)
    model_dt = train_decision_tree(X_train_e, y_train_e, max_depth=3, criterion="gini", min_samples_leaf=10, random_state=42)
    # Huấn luyện mô hình đối chứng Logistic Regression
    model_lr = train_baseline_model(X_train_e, y_train_e, random_state=42)

    # -------------------------------------------------------------
    # GIAI ĐOẠN 4: ĐÁNH GIÁ TRÊN TẬP KIỂM THỬ ĐỘC LẬP (TEST SET)
    # -------------------------------------------------------------
    print("\n[Giai đoạn 4] Đánh giá trên tập kiểm thử độc lập (332 sinh viên)...")
    y_pred_dt = model_dt.predict(X_test_e)
    y_pred_lr = model_lr.predict(X_test_e)

    metrics_dt = calculate_metrics(y_test_e, y_pred_dt)
    metrics_lr = calculate_metrics(y_test_e, y_pred_lr)
    cm_dt = confusion_matrix(y_test_e, y_pred_dt)

    print("\n" + "=" * 55)
    print(" BẢNG ĐỐI SÁNH HIỆU QUẢ PHÂN LỚP (TẬP KIỂM THỬ 20%)")
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

    # Phân tích định lượng ca dự đoán sai (Near-Threshold Error Analysis)
    err_stats = analyze_errors_quantitatively(model_dt, X_test_e, y_test_e, df)
    print(f"\n -> Phân tích định lượng lỗi phân lớp:")
    print(f"    * Tổng số ca sai: {err_stats['total_errors']} / {len(X_test_e)} ({err_stats['error_rate']}%)")
    print(f"    * False Positive: {err_stats['false_positives']} ca (GPA trung vị: {err_stats['fp_gpa_median']}, trong đó {err_stats['fp_near_threshold']} ca có GPA >= 2.30 - đạt {err_stats['fp_near_pct']}%)")
    print(f"    * False Negative: {err_stats['false_negatives']} ca (GPA trung vị: {err_stats['fn_gpa_median']}, trong đó {err_stats['fn_near_threshold']} ca có GPA < 2.70 - đạt {err_stats['fn_near_pct']}%)")
    print(f"    * Tổng ca sai ở vùng cận biên [2.30, 2.70): {err_stats['total_near_threshold']} / {err_stats['total_errors']} ({err_stats['near_threshold_pct']}%)")

    # -------------------------------------------------------------
    # GIAI ĐOẠN 5: LƯU TRỮ METRICS CHO BÁO CÁO
    # -------------------------------------------------------------
    metrics_payload = {
        "dataset_name": "Hồ sơ học tập sinh viên đại học (Anonymized Academic Benchmark)",
        "summary": summary,
        "scenario_early_decision_tree": metrics_dt,
        "scenario_early_logistic_regression": metrics_lr,
        "depth_tuning_cv": depth_df.to_dict(orient="records"),
        "top_features": df_fi.head(8).to_dict(orient="records"),
        "error_analysis": err_stats
    }

    metrics_file = os.path.join(outputs_dir, "metrics.json")
    with open(metrics_file, "w", encoding="utf-8") as f:
        json.dump(metrics_payload, f, ensure_ascii=False, indent=4)

    print(f"\n[Hoàn thành] Đã cập nhật metrics.json tại: {metrics_file}")
    print("=" * 75)


if __name__ == "__main__":
    main()
