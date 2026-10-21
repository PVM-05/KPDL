"""
Module: data_loader.py
Nhiệm vụ: Đọc và khám phá dữ liệu Hồ sơ học tập sinh viên đại học — Đại học Công nghiệp Hà Nội (HaUI).
"""

import os
from typing import Dict, Any, Tuple
import pandas as pd


def load_student_data(data_dir: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Đọc dữ liệu hồ sơ sinh viên HaUI từ file Excel labeled và từ điển biến.
    
    Trả về:
    - df: DataFrame chứa 1.656 sinh viên và 56 biến (kèm tên môn học tường minh).
    - df_dict: DataFrame từ điển giải thích các biến và kỳ học tương ứng.
    """
    file_labeled = os.path.join(data_dir, "01_processed_dataset_labeled.xlsx")
    file_dict = os.path.join(data_dir, "04_variable_dictionary_and_summary.xlsx")
    
    if not os.path.exists(file_labeled):
        raise FileNotFoundError(f"Không tìm thấy file dữ liệu tại: {file_labeled}")
        
    df = pd.read_excel(file_labeled)
    
    df_dict = pd.DataFrame()
    if os.path.exists(file_dict):
        xl = pd.ExcelFile(file_dict)
        if "Variable dictionary" in xl.sheet_names:
            df_dict = pd.read_excel(xl, sheet_name="Variable dictionary")
            
    return df, df_dict


def get_dataset_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Trích xuất các thông số tóm tắt về kích thước, phân bố điểm GPA và thông tin nhân khẩu học.
    """
    n_students = len(df)
    n_features = df.shape[1]
    
    gpa_series = df["GPA"] if "GPA" in df.columns else pd.Series()
    
    summary = {
        "n_students": n_students,
        "n_features": n_features,
        "gpa_mean": float(gpa_series.mean()) if not gpa_series.empty else 0.0,
        "gpa_std": float(gpa_series.std()) if not gpa_series.empty else 0.0,
        "gpa_median": float(gpa_series.median()) if not gpa_series.empty else 0.0,
        "gpa_min": float(gpa_series.min()) if not gpa_series.empty else 0.0,
        "gpa_max": float(gpa_series.max()) if not gpa_series.empty else 0.0,
        "birthplace_count": int(df["BirthPlace"].nunique()) if "BirthPlace" in df.columns else 0,
        "male_count": int((df["Sex"] == 0).sum()) if "Sex" in df.columns else 0,
        "female_count": int((df["Sex"] == 1).sum()) if "Sex" in df.columns else 0,
    }
    return summary
