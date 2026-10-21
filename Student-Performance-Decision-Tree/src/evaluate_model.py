"""
Module: evaluate_model.py
Nhiệm vụ: Đánh giá mô hình Decision Tree trên dữ liệu sinh viên và xuất 5 biểu đồ báo cáo.
Tích hợp phân tích lỗi cận biên định lượng (Near-Threshold Error Analysis).
"""

import os
from typing import Dict, Any, List
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    confusion_matrix
)
from sklearn.tree import plot_tree, DecisionTreeClassifier

# Thiết lập định dạng biểu đồ khoa học
plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")
plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial"]
plt.rcParams["axes.unicode_minus"] = False


def calculate_metrics(y_true, y_pred) -> Dict[str, float]:
    """
    Tính toán các chỉ số đánh giá phân lớp nhị phân:
    Accuracy, Precision, Recall, F1-score (cho lớp Khá/Giỏi và Trung bình/Yếu).
    """
    cm = confusion_matrix(y_true, y_pred)
    tn, fp, fn, tp = cm.ravel()
    
    return {
        "accuracy": float(accuracy_score(y_true, y_pred)),
        "precision_pass": float(precision_score(y_true, y_pred, pos_label=1, zero_division=0)),
        "recall_pass": float(recall_score(y_true, y_pred, pos_label=1, zero_division=0)),
        "f1_pass": float(f1_score(y_true, y_pred, pos_label=1, zero_division=0)),
        "precision_fail": float(precision_score(y_true, y_pred, pos_label=0, zero_division=0)),
        "recall_fail": float(recall_score(y_true, y_pred, pos_label=0, zero_division=0)),
        "f1_fail": float(f1_score(y_true, y_pred, pos_label=0, zero_division=0)),
        "macro_f1": float(f1_score(y_true, y_pred, average="macro")),
        "tn": int(tn),
        "fp": int(fp),
        "fn": int(fn),
        "tp": int(tp),
    }


def plot_class_distribution(df: pd.DataFrame, output_path: str):
    """
    Biểu đồ 1: Phân bố xếp loại học lực của 1.656 sinh viên.
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    pass_count = (df["GPA"] >= 2.50).sum()
    fail_count = (df["GPA"] < 2.50).sum()
    total = len(df)
    
    fig, ax = plt.subplots(figsize=(7.5, 5), dpi=300)
    bars = ax.bar(
        ["Trung bình - Yếu (GPA < 2.5)", "Khá - Giỏi (GPA >= 2.5)"],
        [fail_count, pass_count],
        color=["#e67e22", "#2980b9"],
        edgecolor="#2c3e50",
        width=0.45
    )
    
    for bar in bars:
        h = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2, 
            h + 20, 
            f"{h} sinh viên\n({h / total * 100:.2f}%)", 
            ha="center", va="bottom", fontsize=10.5, fontweight="bold"
        )
        
    ax.set_title("Phân Bố Xếp Loại Học Lực Sinh Viên (N = 1.656)", fontsize=13, fontweight="bold", pad=15)
    ax.set_ylabel("Số lượng sinh viên", fontsize=11)
    ax.set_ylim(0, total * 0.70)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()


def plot_correlation_heatmap(df: pd.DataFrame, output_path: str):
    """
    Biểu đồ 2: Ma trận tương quan giữa GPA và các môn học nền tảng Năm 1.
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    key_cols = [
        "GPA",
        "Advanced Mathematics 1 (Calculus I)", "Advanced Mathematics 2 (Calculus II)",
        "Physics 1", "Theoretical Mechanics", "Strength of Materials",
        "Automotive Electrical and Electronics", "Materials Science",
        "English 1", "English 2"
    ]
    cols_present = [c for c in key_cols if c in df.columns]
    
    corr = df[cols_present].corr()
    
    short_names = {
        "Advanced Mathematics 1 (Calculus I)": "Toán 1 (Giải tích 1)",
        "Advanced Mathematics 2 (Calculus II)": "Toán 2 (Giải tích 2)",
        "Physics 1": "Vật lý 1",
        "Theoretical Mechanics": "Cơ lý thuyết",
        "Strength of Materials": "Sức bền vật liệu",
        "Automotive Electrical and Electronics": "Điện - Điện tử ô tô",
        "Materials Science": "Vật liệu học",
        "English 1": "Tiếng Anh 1",
        "English 2": "Tiếng Anh 2"
    }
    corr_renamed = corr.rename(index=short_names, columns=short_names)
    
    plt.figure(figsize=(9.5, 7.5), dpi=300)
    sns.heatmap(
        corr_renamed, 
        annot=True, 
        fmt=".2f", 
        cmap="Blues", 
        vmin=-0.1, vmax=1.0, 
        linewidths=0.5,
        square=True,
        cbar_kws={"shrink": 0.8}
    )
    plt.title("Ma Trận Tương Quan Giữa GPA Và Các Môn Học Năm 1", fontsize=13, fontweight="bold", pad=15)
    plt.xticks(rotation=45, ha="right", fontsize=9.5)
    plt.yticks(fontsize=9.5)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()


def plot_decision_tree_graph(
    model: DecisionTreeClassifier, 
    feature_names: List[str], 
    output_path: str
):
    """
    Biểu đồ 3: Cấu trúc Cây quyết định dự đoán kết quả học tập.
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    clean_feature_names = [
        f.replace("Advanced Mathematics 1 (Calculus I)", "Toán 1 (Giải tích 1)")
         .replace("Automotive Electrical and Electronics", "Điện - ĐT ô tô")
         .replace("Theoretical Mechanics", "Cơ lý thuyết")
         .replace("Materials Science", "Vật liệu học")
         .replace("Strength of Materials", "Sức bền VL")
         .replace("English 2", "Tiếng Anh 2")
         .replace("English 1", "Tiếng Anh 1")
         .replace("Engineering Thermodynamics", "Nhiệt KT")
         .replace("Advanced Mathematics 2 (Calculus II)", "Toán 2 (Giải tích 2)")
         .replace("National Defense and Security Education", "GD Quốc phòng")
         .replace("Descriptive Geometry", "Hình họa")
         .replace("General Chemistry", "Hóa đại cương")
         .replace("Theory of Machines", "Nguyên lý máy")
         .replace("Engineering Drawing", "Vẽ kỹ thuật")
        for f in feature_names
    ]
    
    fig, ax = plt.subplots(figsize=(18, 9), dpi=300)
    plot_tree(
        model,
        feature_names=clean_feature_names,
        class_names=["Trung bình-Yếu", "Khá-Giỏi"],
        filled=True,
        rounded=True,
        fontsize=10,
        ax=ax,
        precision=2
    )
    ax.set_title("Cấu Trúc Cây Quyết Định Dự Đoán Sớm Kết Quả Học Tập", fontsize=14, fontweight="bold", pad=15)
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()


def plot_confusion_matrix_heatmap(
    cm: np.ndarray, 
    output_path: str
):
    """
    Biểu đồ 4: Ma trận nhầm lẫn (Confusion Matrix).
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    plt.figure(figsize=(6, 5), dpi=300)
    labels = ["Trung bình - Yếu (0)", "Khá - Giỏi (1)"]
    sns.heatmap(
        cm, 
        annot=True, 
        fmt="d", 
        cmap="Blues", 
        xticklabels=labels, 
        yticklabels=labels,
        cbar=False,
        annot_kws={"size": 13, "weight": "bold"}
    )
    plt.title("Ma Trận Nhầm Lẫn (Confusion Matrix)", fontsize=13, fontweight="bold", pad=15)
    plt.xlabel("Nhãn Dự Đoán (Predicted)", fontsize=11, fontweight="bold")
    plt.ylabel("Nhãn Thực Tế (Actual)", fontsize=11, fontweight="bold")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()


def plot_feature_importance_bar(
    model: DecisionTreeClassifier, 
    feature_names: List[str], 
    output_path: str,
    top_n: int = 10
):
    """
    Biểu đồ 5: Mức độ quan trọng của các thuộc tính (Gini Importance).
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    importances = model.feature_importances_
    
    clean_names = [
        f.replace("Automotive Electrical and Electronics", "Điện - Điện tử ô tô")
         .replace("Materials Science", "Vật liệu học")
         .replace("English 2", "Tiếng Anh 2")
         .replace("Engineering Thermodynamics", "Nhiệt kỹ thuật")
         .replace("English 1", "Tiếng Anh 1")
         .replace("Theoretical Mechanics", "Cơ lý thuyết")
         .replace("Advanced Mathematics 1 (Calculus I)", "Toán 1 (Giải tích 1)")
         .replace("Strength of Materials", "Sức bền vật liệu")
        for f in feature_names
    ]
    
    df_imp = pd.DataFrame({
        "feature": clean_names,
        "importance": importances
    }).sort_values("importance", ascending=False).reset_index(drop=True)
    
    top_df = df_imp.head(top_n).sort_values("importance", ascending=True)
    
    plt.figure(figsize=(9.5, 6), dpi=300)
    bars = plt.barh(top_df["feature"], top_df["importance"], color="#2980b9", edgecolor="#1a5276")
    
    for bar in bars:
        w = bar.get_width()
        if w > 0:
            plt.text(w + 0.005, bar.get_y() + bar.get_height() / 2, f"{w:.3f}", 
                     va="center", ha="left", fontsize=9.5, fontweight="bold")
            
    plt.title(f"Top {top_n} Môn Học Năm 1 Có Ảnh Hưởng Lớn Nhất (Gini Importance)", fontsize=13, fontweight="bold", pad=15)
    plt.xlabel("Trọng số phân tách (Feature Importance)", fontsize=11)
    plt.xlim(0, max(top_df["importance"].max() * 1.15, 0.1))
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    
    return df_imp


def analyze_errors_quantitatively(
    model: DecisionTreeClassifier, 
    X_test: pd.DataFrame, 
    y_test: pd.Series,
    original_df: pd.DataFrame
) -> Dict[str, Any]:
    """
    Phân tích định lượng các trường hợp dự đoán sai:
    Tính toán phân vị GPA của các ca FP và FN để chứng minh giả thuyết vùng cận biên (Near-Threshold).
    """
    y_pred = model.predict(X_test)
    test_indices = X_test.index
    
    mis_mask = (y_test.values != y_pred)
    mis_indices = test_indices[mis_mask]
    
    df_errors = original_df.loc[mis_indices].copy()
    df_errors["Actual_GPA"] = original_df.loc[mis_indices, "GPA"]
    df_errors["Actual_Label"] = y_test.loc[mis_indices].values
    df_errors["Predicted_Label"] = y_pred[mis_mask]
    
    total_errors = len(df_errors)
    total_test = len(X_test)
    error_rate = round(float(total_errors / total_test * 100), 2)
    
    fp_mask = (df_errors["Actual_Label"] == 0) & (df_errors["Predicted_Label"] == 1)
    fn_mask = (df_errors["Actual_Label"] == 1) & (df_errors["Predicted_Label"] == 0)
    
    df_fp = df_errors[fp_mask]
    df_fn = df_errors[fn_mask]
    
    fp_count = int(len(df_fp))
    fn_count = int(len(df_fn))
    
    # Số ca FP có GPA >= 2.30 (sát ngưỡng 2.50)
    fp_near_threshold = int((df_fp["Actual_GPA"] >= 2.30).sum()) if fp_count > 0 else 0
    fp_near_pct = round(float(fp_near_threshold / fp_count * 100), 2) if fp_count > 0 else 0.0
    
    # Số ca FN có GPA < 2.70 (sát ngưỡng 2.50)
    fn_near_threshold = int((df_fn["Actual_GPA"] < 2.70).sum()) if fn_count > 0 else 0
    fn_near_pct = round(float(fn_near_threshold / fn_count * 100), 2) if fn_count > 0 else 0.0
    
    # Tổng số ca sai trong vùng cận biên [2.30, 2.70)
    total_near_threshold = fp_near_threshold + fn_near_threshold
    near_threshold_pct = round(float(total_near_threshold / total_errors * 100), 2) if total_errors > 0 else 0.0
    
    return {
        "total_errors": total_errors,
        "error_rate": error_rate,
        "false_positives": fp_count,
        "false_negatives": fn_count,
        "fp_gpa_mean": round(float(df_fp["Actual_GPA"].mean()), 3) if fp_count > 0 else 0.0,
        "fp_gpa_median": round(float(df_fp["Actual_GPA"].median()), 3) if fp_count > 0 else 0.0,
        "fp_near_threshold": fp_near_threshold,
        "fp_near_pct": fp_near_pct,
        "fn_gpa_mean": round(float(df_fn["Actual_GPA"].mean()), 3) if fn_count > 0 else 0.0,
        "fn_gpa_median": round(float(df_fn["Actual_GPA"].median()), 3) if fn_count > 0 else 0.0,
        "fn_near_threshold": fn_near_threshold,
        "fn_near_pct": fn_near_pct,
        "total_near_threshold": total_near_threshold,
        "near_threshold_pct": near_threshold_pct,
    }
