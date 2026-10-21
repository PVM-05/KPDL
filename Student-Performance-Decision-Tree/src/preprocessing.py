"""
Module: preprocessing.py
Nhiệm vụ: Tiền xử lý dữ liệu hồ sơ sinh viên HaUI, xác định nhãn học lực và chia tập huấn luyện/kiểm thử.
"""

from typing import Tuple, List, Optional
import pandas as pd
from sklearn.model_selection import train_test_split


def create_target(df: pd.DataFrame, threshold: float = 2.50) -> pd.Series:
    """
    Chuyển đổi điểm tích lũy toàn khóa GPA (thang 4.0) thành bài toán phân lớp:
    - 1 (Khá - Giỏi): GPA >= 2.50 (chuẩn tốt nghiệp loại Khá trở lên theo quy chế đào tạo tín chỉ)
    - 0 (Trung bình - Yếu): GPA < 2.50 (sinh viên có kết quả học tập ở mức trung bình hoặc có nguy cơ cảnh báo học vụ)
    """
    return (df["GPA"] >= threshold).astype(int)


def clean_course_grades(df: pd.DataFrame, df_dict: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """
    Điền giá trị 0 cho các ô điểm khuyết thiếu (NaN) trong các môn học.
    Theo ghi chú kỹ thuật của bộ dữ liệu HaUI, giá trị 0 đại diện cho việc môn học
    chưa được tích lũy hoặc sinh viên vắng thi/bị điểm F.
    """
    df_clean = df.copy()
    
    # Xác định danh sách các cột môn học
    if df_dict is not None and not df_dict.empty and "Readable meaning" in df_dict.columns:
        course_cols = df_dict[df_dict["Category"] == "Harmonized course-grade variable"]["Readable meaning"].tolist()
        course_cols = [c for c in course_cols if c in df_clean.columns]
    else:
        # Nhận diện tự động các cột không phải metadata
        exclude_cols = ["BirthPlace", "Sex", "GPA", "ScholarshipRecipientsCount", "DisciplineForm", "DisciplineTypeName"]
        course_cols = [c for c in df_clean.columns if c not in exclude_cols]
        
    for col in course_cols:
        df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0.0)
        
    return df_clean


def prepare_features(
    df: pd.DataFrame, 
    df_dict: Optional[pd.DataFrame] = None,
    scenario: str = "early"
) -> Tuple[pd.DataFrame, pd.Series]:
    """
    Chuẩn bị ma trận đặc trưng X và nhãn mục tiêu y theo 2 kịch bản thực nghiệm:
    
    1. scenario="early" (Dự đoán sớm - Early Prediction):
       Chỉ sử dụng 15 môn học của Năm thứ nhất (Học kỳ 1 & 2) để phát hiện sớm
       xu hướng học tập của sinh viên ngay sau năm 1.
       
    2. scenario="full" (Toàn khóa học - Full Profile):
       Sử dụng toàn bộ 50 môn học kết hợp với chỉ số Số lần nhận học bổng.
    """
    df_clean = clean_course_grades(df, df_dict)
    y = create_target(df_clean, threshold=2.50)
    
    # Lấy danh sách môn học Năm 1 từ từ điển nếu có
    if scenario == "early":
        if df_dict is not None and not df_dict.empty and "Semester suffix" in df_dict.columns:
            sem1_2_vars = df_dict[df_dict["Semester suffix"].isin([1.0, 2.0])]["Readable meaning"].tolist()
            selected_features = [c for c in sem1_2_vars if c in df_clean.columns]
        else:
            # Fallback danh sách 15 môn Năm 1 chuẩn của HaUI
            selected_features = [
                "Theoretical Mechanics", "National Defense and Security Education",
                "Descriptive Geometry", "General Chemistry", "English 1",
                "Advanced Mathematics 1 (Calculus I)", "Physics 1", "Theory of Machines",
                "Automotive Electrical and Electronics", "Engineering Thermodynamics",
                "Strength of Materials", "English 2",
                "Advanced Mathematics 2 (Calculus II)", "Materials Science", "Engineering Drawing"
            ]
            selected_features = [c for c in selected_features if c in df_clean.columns]
    else:
        # Kịch bản full: lấy tất cả các môn học và số lần nhận học bổng
        exclude = ["BirthPlace", "Sex", "GPA", "DisciplineForm", "DisciplineTypeName"]
        selected_features = [c for c in df_clean.columns if c not in exclude]

    X = df_clean[selected_features].copy()
    return X, y


def split_data(
    X: pd.DataFrame, 
    y: pd.Series, 
    test_size: float = 0.2, 
    random_state: int = 42
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.Series, pd.Series]:
    """
    Phân chia dữ liệu thành tập Huấn luyện (Train 80%) và Kiểm thử (Test 20%) 
    có phân tầng (stratify=y) để bảo toàn tỷ lệ Khá-Giỏi / Trung bình-Yếu.
    """
    return train_test_split(
        X, y, 
        test_size=test_size, 
        random_state=random_state, 
        stratify=y
    )
