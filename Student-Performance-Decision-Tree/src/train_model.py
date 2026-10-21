"""
Module: train_model.py
Nhiệm vụ: Khởi tạo và huấn luyện mô hình Cây quyết định (Decision Tree) và mô hình đối chứng.
"""

from typing import Dict, Any, List
import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression


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
    với ràng buộc độ sâu (max_depth) và số mẫu tối thiểu ở nút lá (min_samples_leaf)
    nhằm chống hiện tượng quá khớp (Overfitting) và tạo ra cây dễ giải thích.
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
    để so sánh hiệu quả phân loại với mô hình Cây quyết định.
    """
    lr = LogisticRegression(max_iter=1000, random_state=random_state)
    lr.fit(X_train, y_train)
    return lr


def tune_tree_depth(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    max_depth_range=range(1, 9)
) -> pd.DataFrame:
    """
    Khảo sát ảnh hưởng của độ sâu cây (max_depth) đến độ chính xác trên tập Train và Test.
    Dùng để phát hiện điểm cân bằng tối ưu giữa Underfitting và Overfitting.
    """
    results = []
    for d in max_depth_range:
        clf = DecisionTreeClassifier(max_depth=d, min_samples_leaf=10, random_state=42)
        clf.fit(X_train, y_train)
        train_acc = clf.score(X_train, y_train)
        test_acc = clf.score(X_test, y_test)
        results.append({
            "max_depth": d,
            "train_accuracy": round(train_acc, 4),
            "test_accuracy": round(test_acc, 4),
            "generalization_gap": round(train_acc - test_acc, 4)
        })
    return pd.DataFrame(results)
