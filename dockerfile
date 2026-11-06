# استخدام صورة بايثون الرسمية الخفيفة
FROM python:3.10-slim

# تعيين مجلد العمل داخل الحاوية
WORKDIR /app

# تثبيت متطلبات النظام الأساسية إذا لزم الأمر
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    software-properties-common \
    git \
    && rm -rf /var/lib/apt/lists/*

# نسخ ملف المتطلبات أولاً للاستفادة من التخزين المؤقت (Caching)
COPY requirements.txt .

# تثبيت مكتبات بايثون
RUN pip install --no-cache-dir -r requirements.txt

# نسخ باقي ملفات المشروع إلى الحاوية
COPY . .

# إنشاء مستخدم غير ممتاز (Non-root user) لمتطلبات الأمان في Hugging Face
RUN useradd -m -u 1000 user
USER user
ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH

# تعيين مجلد العمل للمستخدم الجديد
WORKDIR /home/user/app
COPY --chown=user:user . /home/user/app

# فتح المنفذ (Port) الاقتصادي لـ Streamlit
EXPOSE 7860

# أمر تشغيل التطبيق (مع ربطه بالمنفذ 7860 والعنوان 0.0.0.0)
CMD ["streamlit", "run", "app/etabs_tables.py", "--server.port=7860", "--server.address=0.0.0.0"]