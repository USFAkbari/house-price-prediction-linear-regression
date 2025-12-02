#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اسکریپت برای استخراج پارامترهای مدل و ذخیره در فایل JSON
"""

import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer
import json
import sys

# خواندن فایل اکسل
print("در حال خواندن فایل اکسل...")
try:
    df = pd.read_excel('تمرین سوم.xlsx')
except Exception as e:
    print(f"خطا در خواندن فایل: {e}")
    sys.exit(1)

# پاک‌سازی داده‌ها
for col in df.columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

df_clean = df.copy()
df_clean = df_clean.dropna(thresh=len(df_clean.columns) // 2)
df = df_clean

# شناسایی ستون price
if 'price' in df.columns:
    price_col = 'price'
elif 'Price' in df.columns:
    price_col = 'Price'
elif 'PRICE' in df.columns:
    price_col = 'PRICE'
else:
    price_col = None
    for col in df.columns:
        if 'قیمت' in str(col) or 'price' in str(col).lower():
            price_col = col
            break
    
    if price_col is None:
        print("\nخطا: ستون price یافت نشد!")
        sys.exit(1)

# جداسازی ویژگی‌ها و قیمت
feature_cols = [col for col in df.columns if col != price_col]

# جداسازی داده‌های آموزشی و تست
train_data = df[df[price_col].notna()].copy()
train_data = train_data.dropna(subset=feature_cols, how='all')

# آماده‌سازی داده‌ها
X_train = train_data[feature_cols].copy()
y_train = train_data[price_col].copy()

for col in feature_cols:
    X_train[col] = pd.to_numeric(X_train[col], errors='coerce')
y_train = pd.to_numeric(y_train, errors='coerce')

X_train = X_train.values
y_train = y_train.values

# بررسی مقادیر خالی
if pd.DataFrame(X_train, columns=feature_cols).isnull().sum().sum() > 0:
    imputer = SimpleImputer(strategy='mean')
    X_train = imputer.fit_transform(X_train)

# نرمال‌سازی ویژگی‌ها
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)

# آموزش مدل رگرسیون خطی
model = LinearRegression()
model.fit(X_train_scaled, y_train)

# استخراج پارامترها
coefficients = model.coef_.tolist()
intercept = float(model.intercept_)
means = scaler.mean_.tolist()
stds = scaler.scale_.tolist()

# ذخیره پارامترها در فایل JSON
params = {
    'feature_names': feature_cols,
    'coefficients': coefficients,
    'intercept': intercept,
    'means': means,
    'stds': stds
}

with open('model_params.json', 'w', encoding='utf-8') as f:
    json.dump(params, f, ensure_ascii=False, indent=2)

print("پارامترهای مدل در فایل 'model_params.json' ذخیره شد.")
print(f"\nIntercept: {intercept}")
print(f"\nتعداد ویژگی‌ها: {len(feature_cols)}")
for i, (name, coef, mean, std) in enumerate(zip(feature_cols, coefficients, means, stds)):
    print(f"{i+1}. {name}: coef={coef:.4f}, mean={mean:.4f}, std={std:.4f}")

