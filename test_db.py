import os
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")

print(f"URL we're using: {DATABASE_URL}")

try:
    engine = create_engine(DATABASE_URL)
    with engine.connect() as conn:
        result = conn.execute(text("SELECT 1")).fetchone()
        print("✅ Database connection successful!", result)
except Exception as e:
    print(f"❌ Database connection failed: {e}")