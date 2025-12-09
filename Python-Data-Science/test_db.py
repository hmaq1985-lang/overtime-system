import os
from sqlalchemy import create_engine, text

# قراءة المتغير من Environment
DATABASE_URL = os.environ.get("DATABASE_URL")

# إنشاء الاتصال
engine = create_engine(DATABASE_URL)

try:
    with engine.connect() as conn:
        result = conn.execute(text("SELECT 1"))
        print("✅ الاتصال ناجح! قاعدة البيانات تعمل.")
except Exception as e:
    print("❌ فشل الاتصال:")
    print(e)
