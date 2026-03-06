from __future__ import annotations

import pandas as pd

# 店舗名キーワードから自動分類する辞書
CATEGORY_RULES = {
    "水道光熱費": (
        "東京電力",
        "関西電力",
        "電力",
        "ガス",
        "水道",
        "電気料金",
    ),
    "サブスク（定期サービス）": (
        "netflix",
        "amazon prime",
        "spotify",
        "apple",
        "icloud",
        "google",
        "chatgpt",
        "adobe",
        "microsoft",
    ),
}


def classify_with_fallback(store_name: str, fallback_category: str) -> str:
    if pd.isna(store_name):
        return fallback_category

    text = str(store_name).lower()

    for category_name, keywords in CATEGORY_RULES.items():
        if any(keyword.lower() in text for keyword in keywords):
            return category_name

    return fallback_category
