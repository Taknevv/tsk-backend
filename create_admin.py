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

# Check if admin already exists
existing = db.query(User).filter(User.email == "admin@tsk.com").first()
if existing:
    print("Admin user already exists.")
else:
    admin = User(
        email="admin@tsk.com",
        hashed_password=get_password_hash("admin123"),
        role="ua",
        name="Ultimate Admin"
    )
    db.add(admin)
    db.commit()
    print("✅ Admin user created: admin@tsk.com / admin123")