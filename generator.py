import pandas as pd
import requests
import time
import re

# مسیر فایل‌های ورودی و خروجی
input_excel = "produts-1404-03-09-with-images.xlsx"
output_excel = "products_output_with_content.xlsx"

# کلید API مربوط به OpenAI (GPT-3.5-Turbo)
API_KEY = ""  # ← اینجا کلید API خودت رو بزن

# هدرهای درخواست HTTP
headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

# بررسی فارسی بودن محتوا
def is_persian(text, min_persian_chars=30):
    persian_chars = re.findall(r'[\u0600-\u06FF]', text)
    return len(persian_chars) >= min_persian_chars

# شمارش کلمات
def word_count(text):
    words = re.findall(r'\b\w+\b', text)
    return len(words)
# تمیز کردن متن از کاراکترهای کنترلی و نامجاز برای اکسل
def clean_text_for_excel(text):
    if pd.isna(text):
        return ""
    # حذف کاراکترهای کنترلی
    text = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", "", text)
    # اصلاح تگ‌های h ناقص (مثلاً <h2>...</h>)
    text = re.sub(r"<h(\d)>(.*?)</h>", r"<h\1>\2</h\1>", text)
    return text


# ساخت پرامپت بهینه‌شده با قوانین دقیق
def build_prompt(product_name, product_slug):
    return f"""
شما یک نویسنده‌ی حرفه‌ای و متخصص در تولید محتوای فنی و سئو شده هستید. لطفاً درباره قطعه خودرو با نام: {product_name} متنی رسمی، روان و تخصصی در قالب ۶ بخش زیر بنویسید:

1. معرفی محصول و نقش آن در خودرو
2. ویژگی‌ها و مشخصات فنی
3. علائم خرابی یا نقص عملکرد
4. نکات ایمنی و مراقبتی هنگام استفاده
5. نکات مهم هنگام خرید
6. راهنمای نصب یا تعویض

**دستورالعمل مهم:**

- در **اولین پاراگراف بخش معرفی محصول**، جمله‌ی زیر را اضافه کن:  
برای خرید <a href="/{product_slug}" title="{product_name}">{product_name}</a> اینجا کلیک کنید.

- در انتهای محتوا یک پاراگراف اضافه کن که شامل این لینک باشد:  
برای <a href="https://saeinkala.ir/shop">مشاهده سایر لوازم و قطعات خودرو</a> کلیک کنید.

- تمام تیترهای هر بخش باید با تگ HTML <h2> نوشته شود.  
- از ** یا هر علامت دیگر به جای h2 استفاده نکن.

- کل متن بین ۳۰۰ تا ۴۰۰ کلمه باشد.  
→ اگر اطلاعات کافی درباره محصول وجود ندارد، از نوشتن جملات غیرمفید و بی‌ربط خودداری کن.  
→ فقط مطالب مرتبط و باکیفیت بنویس.

- خروجی باید فقط به زبان فارسی و بدون کلمات خارجی غیرضروری باشد.  
- خروجی به صورت HTML ساده باشد و فقط از تگ‌های: <h2>, <p>, <ul>, <li>, <a> استفاده کن.
"""


# ارسال درخواست به API و دریافت محتوا
def generate_content(product_name, product_slug, max_attempts=5):
    prompt = build_prompt(product_name, product_slug)
    payload = {
        "model": "gpt-3.5-turbo",
        "messages": [
            {"role": "system", "content": "شما فقط متن فارسی تولید می‌کنید."},
            {"role": "user", "content": prompt}
        ]
    }

    for attempt in range(max_attempts):
        try:
            response = requests.post(
                "https://api.openai.com/v1/chat/completions",
                json=payload,
                headers=headers,
                timeout=30
            )
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"].strip()

            if not is_persian(content):
                print(f"❌ تلاش {attempt+1}: متن غیر فارسی، تلاش مجدد...")
                time.sleep(10)
                continue

            wc = word_count(content)
            if wc < 260 or wc > 600:  # حدود ۳۵۰ کلمه ± رنج مجاز
                print(f"❌ تلاش {attempt+1}: متن {wc} کلمه دارد، تلاش مجدد...")
                time.sleep(10)
                continue

            return content

        except Exception as e:
            print(f"[خطا در تلاش {attempt+1}]: {e}")
            time.sleep(15)

    return "[محتوای معتبر تولید نشد]"

# پاک کردن لاگ قبلی (هر بار تازه)
open("failed_products.txt", "w", encoding="utf-8").close()

# خواندن فایل اکسل و تولید محتوا
df = pd.read_excel(input_excel)
df["محتوا"] = ""

for idx, row in df.iterrows():
    product_name = str(row["نام"]).strip()
    product_slug = product_name.strip()

    print(f"⏳ در حال تولید محتوا برای: {product_name} → /{product_slug}")

    content = generate_content(product_name, product_slug)
    df.at[idx, "محتوا"] = content

    # ذخیره هر 10 محصول
    if (idx + 1) % 10 == 0:
        df.to_excel(output_excel, index=False)
        print(f"💾 Progress Saved: {idx + 1} محصول ذخیره شد.")

    # ذخیره محصول مشکل‌دار در لاگ
    if content == "[محتوای معتبر تولید نشد]":
        with open("failed_products.txt", "a", encoding="utf-8") as log_file:
            log_file.write(f"{product_name} → /{product_slug}\n")
        print(f"⚠ محصول '{product_name}' در لاگ ثبت شد.")

    print(f"✔ '{product_name}' انجام شد.\n")
    time.sleep(1.5)



# تمیز کردن ستون محتوا برای اکسل (قبل از ذخیره)
df["محتوا"] = df["محتوا"].apply(clean_text_for_excel)

# ذخیره فایل نهایی
df.to_excel(output_excel, index=False)
print(f"✅ فایل خروجی با موفقیت ذخیره شد: {output_excel}")


