"""
Module: train_model.py
Nhiệm vụ: Khởi tạo và huấn luyện mô hình Cây quyết định (Decision Tree) và mô hình đối chứng.
Áp dụng 5-Fold Stratified Cross-Validation trên tập Train nhằm ngăn chặn rò rỉ dữ liệu (Zero Test Leakage).
"""

from typing import Dict, Any, List
import pandas as pd
import numpy as np
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import cross_val_score, StratifiedKFold


def train_decision_tree(
    X_train: pd.DataFrame, 
    y_train: pd.Series, 
    max_depth: int = 3, 
    criterion: str = "gini",
    min_samples_leaf: int = 10,
    random_state: int = 42
) -> DecisionTreeClassifier:
    """
    Huấn luyện mô hình Cây quyết định (Decision Tree Classifier)
    với cấu hình kiểm soát độ sâu (Pre-pruning) nhằm tối ưu khả năng diễn giải.
    """
    clf = DecisionTreeClassifier(
        criterion=criterion,
        max_depth=max_depth,
        min_samples_leaf=min_samples_leaf,
        random_state=random_state
    )
    clf.fit(X_train, y_train)
    return clf


def train_baseline_model(
    X_train: pd.DataFrame, 
    y_train: pd.Series, 
    random_state: int = 42
) -> LogisticRegression:
    """
    Huấn luyện mô hình đối chứng Hồi quy Logistic (Logistic Regression)
    để so sánh hiệu năng với Cây quyết định trên cùng tập đặc trưng Năm 1.
    """
    lr = LogisticRegression(max_iter=1000, random_state=random_state)
    lr.fit(X_train, y_train)
    return lr


def tune_tree_depth_cv(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    max_depth_range=range(1, 9),
    cv: int = 5,
    random_state: int = 42
) -> pd.DataFrame:
    """
    Khảo sát ảnh hưởng của max_depth bằng 5-Fold Stratified Cross-Validation trên tập Huấn luyện (Train),
    hoàn toàn không sử dụng tập Test để chọn tham số nhằm tránh rò rỉ dữ liệu (Zero Test Leakage).
    Cột Test Accuracy chỉ được ghi nhận song song để đối chiếu hiện tượng Overfitting trong báo cáo.
    """
    skf = StratifiedKFold(n_splits=cv, shuffle=True, random_state=random_state)
    results = []

    for d in max_depth_range:
        clf = DecisionTreeClassifier(
            max_depth=d, 
            min_samples_leaf=10, 
            random_state=random_state
        )
        
        # Đánh giá 5-Fold CV trên tập Train
        cv_scores = cross_val_score(clf, X_train, y_train, cv=skf, scoring="accuracy")
        
        # Huấn luyện trên toàn bộ Train để kiểm tra Test Holdout
        clf.fit(X_train, y_train)
        train_acc = clf.score(X_train, y_train)
        test_acc = clf.score(X_test, y_test)
        
        results.append({
            "max_depth": d,
            "cv_mean_accuracy": round(float(cv_scores.mean()), 4),
            "cv_std_accuracy": round(float(cv_scores.std()), 4),
            "train_accuracy": round(float(train_acc), 4),
            "test_accuracy": round(float(test_acc), 4),
            "generalization_gap": round(float(train_acc - test_acc), 4)
        })

    return pd.DataFrame(results)
