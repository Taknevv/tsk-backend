from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from models import User
from auth import get_password_hash
import os
from dotenv import load_dotenv

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
db = SessionLocal()

user = db.query(User).filter(User.email == "admin@tsk.com").first()
if user:
    user.hashed_password = get_password_hash("admin123")
    db.commit()
    print("✅ Password hash updated")
else:
    print("❌ Admin user not found")