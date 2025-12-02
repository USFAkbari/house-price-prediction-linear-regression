#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
اسکریپت برای محاسبه قیمت خانه‌ها و تحلیل ویژگی‌ها
"""

import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import r2_score
import sys

# خواندن فایل اکسل
print("در حال خواندن فایل اکسل...")
try:
    df = pd.read_excel('تمرین سوم.xlsx')
except Exception as e:
    print(f"خطا در خواندن فایل: {e}")
    sys.exit(1)

print(f"\nشکل داده‌ها: {df.shape}")
print(f"\nنام ستون‌ها:")
print(df.columns.tolist())

# پاک‌سازی داده‌ها: حذف سطرهایی که هدر هستند
print("\nدر حال پاک‌سازی داده‌ها...")
# تبدیل تمام ستون‌ها به عددی (به جز ستون‌هایی که واقعاً عددی نیستند)
for col in df.columns:
    # تبدیل ستون به عددی، سطرهای غیرقابل تبدیل به NaN تبدیل می‌شوند
    df[col] = pd.to_numeric(df[col], errors='coerce')

# حذف سطرهایی که تمام مقادیرشان NaN است یا سطرهایی که نام ستون‌ها را دارند
df_clean = df.copy()
# حذف سطرهایی که بیش از نیمی از مقادیرشان NaN است (احتمالاً هدر)
df_clean = df_clean.dropna(thresh=len(df_clean.columns) // 2)

print(f"\nشکل داده‌ها پس از پاک‌سازی: {df_clean.shape}")
print(f"\nاولین سطرها:")
print(df_clean.head(10))
print(f"\nآخرین سطرها:")
print(df_clean.tail())
print(f"\nاطلاعات داده‌ها:")
print(df_clean.info())
print(f"\nمقادیر خالی:")
print(df_clean.isnull().sum())

df = df_clean

# شناسایی ستون price
if 'price' in df.columns:
    price_col = 'price'
elif 'Price' in df.columns:
    price_col = 'Price'
elif 'PRICE' in df.columns:
    price_col = 'PRICE'
else:
    # جستجوی ستون قیمت با نام فارسی
    price_col = None
    for col in df.columns:
        if 'قیمت' in str(col) or 'price' in str(col).lower():
            price_col = col
            break
    
    if price_col is None:
        print("\nخطا: ستون price یافت نشد!")
        print("ستون‌های موجود:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        sys.exit(1)

print(f"\nستون قیمت شناسایی شد: {price_col}")

# جداسازی ویژگی‌ها و قیمت
feature_cols = [col for col in df.columns if col != price_col]
print(f"\nویژگی‌ها ({len(feature_cols)}):")
for i, col in enumerate(feature_cols, 1):
    print(f"{i}. {col}")

# بررسی داده‌ها
print(f"\nتعداد کل سطرها: {len(df)}")
print(f"تعداد سطرهای با قیمت: {df[price_col].notna().sum()}")
print(f"تعداد سطرهای بدون قیمت: {df[price_col].isna().sum()}")

# جداسازی داده‌های آموزشی و تست
train_data = df[df[price_col].notna()].copy()
test_data = df[df[price_col].isna()].copy()

print(f"\nداده‌های آموزشی: {len(train_data)} سطر")
print(f"داده‌های تست: {len(test_data)} سطر")

if len(train_data) < 2:
    print("خطا: داده‌های آموزشی کافی نیست!")
    sys.exit(1)

if len(test_data) < 2:
    print("هشدار: کمتر از 2 سطر برای پیش‌بینی وجود دارد!")

# اطمینان از اینکه فقط سطرهای معتبر را انتخاب می‌کنیم
# حذف سطرهایی که تمام ویژگی‌هایشان NaN است
train_data = train_data.dropna(subset=feature_cols, how='all')
test_data = test_data.dropna(subset=feature_cols, how='all')

print(f"\nپس از حذف سطرهای نامعتبر:")
print(f"داده‌های آموزشی: {len(train_data)} سطر")
print(f"داده‌های تست: {len(test_data)} سطر")

# آماده‌سازی داده‌ها - تبدیل به عددی
X_train = train_data[feature_cols].copy()
y_train = train_data[price_col].copy()

# تبدیل به عددی
for col in feature_cols:
    X_train[col] = pd.to_numeric(X_train[col], errors='coerce')
y_train = pd.to_numeric(y_train, errors='coerce')

X_test = test_data[feature_cols].copy()
for col in feature_cols:
    X_test[col] = pd.to_numeric(X_test[col], errors='coerce')

# تبدیل به numpy array
X_train = X_train.values
y_train = y_train.values
X_test = X_test.values

# بررسی مقادیر خالی در ویژگی‌ها
print(f"\nمقادیر خالی در داده‌های آموزشی:")
missing_train = pd.DataFrame(X_train, columns=feature_cols).isnull().sum()
print(missing_train[missing_train > 0])

if missing_train.sum() > 0:
    print("هشدار: مقادیر خالی در ویژگی‌ها وجود دارد. در حال پر کردن با میانگین...")
    from sklearn.impute import SimpleImputer
    imputer = SimpleImputer(strategy='mean')
    X_train = imputer.fit_transform(X_train)
    X_test = imputer.transform(X_test)

# نرمال‌سازی ویژگی‌ها
print("\nدر حال نرمال‌سازی ویژگی‌ها...")
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# آموزش مدل رگرسیون خطی
print("\nدر حال آموزش مدل رگرسیون خطی...")
model = LinearRegression()
model.fit(X_train_scaled, y_train)

# ارزیابی مدل
y_train_pred = model.predict(X_train_scaled)
r2 = r2_score(y_train, y_train_pred)
print(f"\nR² Score روی داده‌های آموزشی: {r2:.4f}")

# پیش‌بینی قیمت برای دو سطر آخر
print("\nدر حال پیش‌بینی قیمت برای سطرهای بدون قیمت...")
y_test_pred = model.predict(X_test_scaled)

# نمایش نتایج
print("\n" + "="*80)
print("نتایج پیش‌بینی قیمت:")
print("="*80)

for i, (idx, row) in enumerate(test_data.iterrows(), 1):
    print(f"\nخانه {i} (سطر {idx + 1}):")
    print(f"  قیمت پیش‌بینی شده: {y_test_pred[i-1]:,.2f}")
    print(f"  ویژگی‌ها:")
    for col in feature_cols:
        print(f"    - {col}: {row[col]}")

# محاسبه درصد تأثیر ویژگی‌ها
print("\n" + "="*80)
print("تحلیل تأثیر ویژگی‌ها بر قیمت:")
print("="*80)

# محاسبه اهمیت نسبی بر اساس ضرایب
coefficients = model.coef_
feature_importance = np.abs(coefficients)
feature_importance_percent = (feature_importance / feature_importance.sum()) * 100

# ایجاد جدول نتایج
importance_df = pd.DataFrame({
    'ویژگی': feature_cols,
    'ضریب': coefficients,
    'اهمیت مطلق': feature_importance,
    'درصد تأثیر': feature_importance_percent
})

# مرتب‌سازی بر اساس درصد تأثیر
importance_df = importance_df.sort_values('درصد تأثیر', ascending=False)

print("\nجدول درصد تأثیر ویژگی‌ها:")
print("-" * 80)
print(importance_df.to_string(index=False))
print("-" * 80)

# توضیح منطقی قیمت‌ها
print("\n" + "="*80)
print("توضیح منطقی قیمت‌های پیش‌بینی شده:")
print("="*80)

# محاسبه میانگین مقادیر برای مقایسه
mean_values = train_data[feature_cols].mean()

for i, (idx, row) in enumerate(test_data.iterrows(), 1):
    print(f"\n{'='*80}")
    print(f"خانه {i} (سطر {idx + 1}):")
    print(f"{'='*80}")
    predicted_price = y_test_pred[i-1]
    print(f"\nقیمت پیش‌بینی شده: {predicted_price:,.2f}")
    
    # تبدیل مقادیر سطر به عددی
    row_values = []
    row_dict = {}
    for col in feature_cols:
        val = pd.to_numeric(row[col], errors='coerce')
        if pd.isna(val):
            val = 0
        row_values.append(val)
        row_dict[col] = val
    
    row_scaled = scaler.transform([row_values])[0]
    contributions = model.coef_ * row_scaled
    
    # تحلیل ویژگی‌های کلیدی
    print(f"\nتحلیل ویژگی‌های کلیدی:")
    print(f"-" * 80)
    
    # ویژگی‌های مثبت (افزایش قیمت)
    positive_features = []
    negative_features = []
    
    for j, (col, contrib, val) in enumerate(zip(feature_cols, contributions, row_values)):
        mean_val = mean_values[col]
        if contrib > 0.1:
            positive_features.append((col, contrib, val, mean_val))
        elif contrib < -0.1:
            negative_features.append((col, contrib, val, mean_val))
    
    if positive_features:
        print("\nویژگی‌هایی که قیمت را افزایش می‌دهند:")
        for col, contrib, val, mean_val in sorted(positive_features, key=lambda x: x[1], reverse=True):
            diff = val - mean_val
            print(f"  • {col}: {val} (میانگین: {mean_val:.1f}, تفاوت: {diff:+.1f}) → +{contrib:.2f}")
    
    if negative_features:
        print("\nویژگی‌هایی که قیمت را کاهش می‌دهند:")
        for col, contrib, val, mean_val in sorted(negative_features, key=lambda x: x[1]):
            diff = val - mean_val
            print(f"  • {col}: {val} (میانگین: {mean_val:.1f}, تفاوت: {diff:+.1f}) → {contrib:.2f}")
    
    # خلاصه منطقی
    print(f"\nخلاصه منطقی:")
    area = row_dict['Area(m2)']
    rooms = row_dict['Rooms']
    age = row_dict['Age']
    distance = row_dict['Distance-to-center']
    
    print(f"این خانه با مساحت {area} متر مربع و {int(rooms)} اتاق،")
    if age < mean_values['Age']:
        print(f"سن کم ({age} سال) دارد که قیمت را افزایش می‌دهد.")
    else:
        print(f"سن بالایی ({age} سال) دارد که قیمت را کاهش می‌دهد.")
    
    if distance < mean_values['Distance-to-center']:
        print(f"نزدیک به مرکز شهر ({distance} کیلومتر) است که قیمت را افزایش می‌دهد.")
    else:
        print(f"دور از مرکز شهر ({distance} کیلومتر) است که قیمت را کاهش می‌دهد.")
    
    amenities = []
    if row_dict['Parking'] == 1:
        amenities.append("پارکینگ")
    if row_dict['Elevator'] == 1:
        amenities.append("آسانسور")
    if row_dict['Renovated'] == 1:
        amenities.append("نوساز")
    if row_dict['Balcony'] == 1:
        amenities.append("بالکن")
    if row_dict['View'] == 1:
        amenities.append("چشم‌انداز")
    
    if amenities:
        print(f"دارای امکانات: {', '.join(amenities)} که قیمت را افزایش می‌دهد.")
    else:
        print(f"بدون امکانات اضافی که قیمت را کاهش می‌دهد.")
    
    print(f"\nقیمت پایه (intercept): {model.intercept_:.2f}")
    print(f"مجموع تأثیر ویژگی‌ها: {contributions.sum():+.2f}")
    print(f"قیمت نهایی: {predicted_price:,.2f}")

# ذخیره نتایج در فایل اکسل
print("\nدر حال ذخیره نتایج در فایل اکسل...")
output_df = df.copy()

# اضافه کردن قیمت‌های پیش‌بینی شده
pred_idx = 0
for idx in test_data.index:
    output_df.loc[idx, price_col] = y_test_pred[pred_idx]
    pred_idx += 1

# ایجاد شیت جدید برای تحلیل ویژگی‌ها
with pd.ExcelWriter('results.xlsx', engine='openpyxl') as writer:
    output_df.to_excel(writer, sheet_name='نتایج پیش‌بینی', index=False)
    importance_df.to_excel(writer, sheet_name='تحلیل ویژگی‌ها', index=False)

print("نتایج در فایل 'results.xlsx' ذخیره شد.")
print("\n" + "="*80)
print("تحلیل کامل شد!")
print("="*80)

