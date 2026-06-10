import os
import tempfile
import math
import numpy as np
import pandas as pd
from datetime import datetime
from typing import List

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.responses import FileResponse
from jose import jwt, JWTError
from sqlalchemy.orm import Session
from dotenv import load_dotenv
from openpyxl import Workbook

from database import get_db, engine
from excel_styles import (
    build_dashboard, build_coil_detail, build_fatigue_log,
    build_inspector_matrix, build_inspector_calc, build_change_log,
    no_grid
)
from models import Base, User, Coil, Inspection, Inspector
from schemas import UserCreate, UserOut, Token, CoilCreate, CoilOut, InspectionCreate, InspectionOut
from auth import get_password_hash, verify_password, create_access_token, SECRET_KEY, ALGORITHM

load_dotenv()

# Create tables
Base.metadata.create_all(bind=engine)

app = FastAPI(title="TSK Coil Backend")

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="login")

# ---------- Helper: get current user ----------
def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        email: str = payload.get("sub")
        if email is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception
    user = db.query(User).filter(User.email == email).first()
    if user is None:
        raise credentials_exception
    return user

# ---------- Auth endpoints ----------
@app.post("/login", response_model=Token)
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.email == form_data.username).first()
    if not user or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(status_code=400, detail="Incorrect email or password")
    access_token = create_access_token(data={"sub": user.email})
    return {"access_token": access_token, "token_type": "bearer"}

@app.post("/register", response_model=UserOut)
def register(user: UserCreate, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    if current_user.role != 'ua':
        raise HTTPException(status_code=403, detail="Only Ultimate Admin can register users")
    hashed = get_password_hash(user.password)
    db_user = User(
        email=user.email,
        hashed_password=hashed,
        role=user.role,
        line=user.line,
        inspector_id=user.inspector_id,
        name=user.name
    )
    db.add(db_user)
    db.commit()
    db.refresh(db_user)
    return db_user

@app.get("/me", response_model=UserOut)
def get_me(current_user: User = Depends(get_current_user)):
    return current_user

# ---------- Coils ----------
@app.post("/coils", response_model=CoilOut)
def create_coil(coil: CoilCreate, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    if current_user.role == 'inspector':
        raise HTTPException(status_code=403, detail="Not authorized")
    db_coil = Coil(**coil.dict(), created_by=current_user.id)
    db.add(db_coil)
    db.commit()
    db.refresh(db_coil)
    return db_coil

@app.get("/coils", response_model=List[CoilOut])
def read_coils(skip: int = 0, limit: int = 100, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    if current_user.role == 'ua':
        coils = db.query(Coil).offset(skip).limit(limit).all()
    elif current_user.role.endswith('_admin'):
        line = current_user.role.split('_')[0].upper()
        coils = db.query(Coil).filter(Coil.line == line).offset(skip).limit(limit).all()
    else:
        coils = []
    return coils

# ---------- Inspections ----------
@app.post("/inspections", response_model=InspectionOut)
def create_inspection(inspection: InspectionCreate, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    if current_user.role == 'inspector':
        raise HTTPException(status_code=403, detail="Not authorized")
    db_inspection = Inspection(**inspection.dict(), created_by=current_user.id)
    db.add(db_inspection)
    db.commit()
    db.refresh(db_inspection)
    return db_inspection

@app.get("/inspections", response_model=List[InspectionOut])
def read_inspections(skip: int = 0, limit: int = 100, db: Session = Depends(get_db), current_user = Depends(get_current_user)):
    if current_user.role == 'ua':
        inspections = db.query(Inspection).offset(skip).limit(limit).all()
    elif current_user.role.endswith('_admin'):
        line = current_user.role.split('_')[0].upper()
        inspections = db.query(Inspection).join(Coil).filter(Coil.line == line).offset(skip).limit(limit).all()
    else:
        # inspector sees only their own inspections
        inspections = db.query(Inspection).filter(Inspection.inspector_id == current_user.inspector_id).offset(skip).limit(limit).all()
    return inspections

# ---------- Helper to compute line statistics ----------
def compute_line_stats(coils, inspections):
    line_stats = {}
    for line in ["CGL", "CAL", "RCL"]:
        line_coils = [c for c in coils if c.line == line]
        n_coils = len(line_coils)
        if n_coils == 0:
            wl_avg = 0.76
            defects_km_avg = 0.5
            cv = 0.3
            fat_avg = 6.5
        else:
            wls = []
            defects_km = []
            for c in line_coils:
                if c.length_m and c.speed_mps:
                    dur_min = c.length_m / (c.speed_mps * 60)
                    wl = (dur_min * 60) / c.length_m if c.length_m else 0
                    wls.append(wl)
                if c.length_m:
                    defects_km.append(c.defect_count / (c.length_m / 1000))
            wl_avg = np.mean(wls) if wls else 0.76
            defects_km_avg = np.mean(defects_km) if defects_km else 0.5
            cv = np.std(defects_km) / defects_km_avg if defects_km_avg > 0 and len(defects_km) > 1 else 0.3
            # Compute average fatigue from inspections on this line
            line_inspections = [i for i in inspections if i.coil_id in [c.id for c in line_coils]]
            fats = [i.fatigue_score_post for i in line_inspections if i.fatigue_score_post]
            fat_avg = np.mean(fats) if fats else 6.5
        speed = {"CGL": 150, "CAL": 200, "RCL": 250}[line]
        raw = (speed * wl_avg) / 60
        n_base = max(1, math.ceil(raw))
        n_peak = math.ceil(n_base * 1.3)
        line_stats[line] = {
            "wl_avg": round(wl_avg, 4),
            "defects_km_avg": round(defects_km_avg, 4),
            "n_coils": n_coils,
            "speed": speed,
            "n_base": n_base,
            "n_peak": n_peak,
            "cv": round(cv, 4),
            "fat_avg": round(fat_avg, 2)
        }
    return line_stats

# ---------- Export endpoint (styled, role‑based, with AI sheets) ----------
@app.get("/export")
def export_results(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    # Fetch data based on role
    if current_user.role == 'ua':
        coils = db.query(Coil).order_by(Coil.created_at.desc()).all()
        inspections = db.query(Inspection).all()
        inspectors = db.query(Inspector).all()
    elif current_user.role.endswith('_admin'):
        line = current_user.role.split('_')[0].upper()
        coils = db.query(Coil).filter(Coil.line == line).order_by(Coil.created_at.desc()).all()
        inspections = db.query(Inspection).join(Coil).filter(Coil.line == line).all()
        inspectors = db.query(Inspector).all()
    else:  # inspector
        coils = []  # inspectors don't see coils
        inspections = db.query(Inspection).filter(Inspection.inspector_id == current_user.inspector_id).all()
        inspectors = db.query(Inspector).filter(Inspector.inspector_id == current_user.inspector_id).all()

    # Convert to DataFrames for AI engine (optional)
    coils_df = pd.DataFrame([c.__dict__ for c in coils]) if coils else pd.DataFrame()
    inspections_df = pd.DataFrame([i.__dict__ for i in inspections]) if inspections else pd.DataFrame()
    inspectors_df = pd.DataFrame([ins.__dict__ for ins in inspectors]) if inspectors else pd.DataFrame()

    # Compute line statistics (respects filtered coils)
    line_stats = compute_line_stats(coils, inspections)

    # Create workbook and remove default sheet
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    #no_grid(wb.active)   # optional

    # Build styled sheets
    build_dashboard(wb, coils, inspections, inspectors, line_stats)
    build_coil_detail(wb, coils, inspections)
    build_fatigue_log(wb, inspections, coils)
    build_inspector_matrix(wb, inspectors)
    build_inspector_calc(wb, line_stats)
    build_change_log(wb, len(coils))

    # Add AI sheets (A1‑A10) using your existing engine (if available)
    try:
        from ai_engine import add_ai_sheets_to_workbook
        wb = add_ai_sheets_to_workbook(wb, coils_df, inspections_df, inspectors_df, line_stats, avail_insp=len(inspectors))
    except Exception as e:
        # If AI sheets fail, add an error sheet for debugging (optional)
        err_ws = wb.create_sheet("AI Error")
        err_ws.append(["Error", str(e)])
        print(f"AI sheets error: {e}")

    # Save to temporary file and return
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    return FileResponse(
        path=tmp_path,
        filename=f"TSK_Final_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- Filtered AI JSON endpoints for Flutter app ----------
@app.get("/api/a1/forecast")
def get_a1_forecast(current_user: User = Depends(get_current_user)):
    """Global 24‑hour demand forecast (for UA)"""
    forecast = [{"hour": i, "defects": round(i * 0.5, 2)} for i in range(1, 25)]
    return forecast

@app.get("/api/a1/forecast/line")
def get_a1_forecast_by_line(line: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    """Return forecast for a specific line. Accessible only to UA or line admin of that line."""
    if current_user.role != 'ua' and current_user.line != line:
        raise HTTPException(status_code=403, detail="Not authorized for this line")
    coils = db.query(Coil).filter(Coil.line == line).all()
    if not coils:
        defects_km = [0.5] * 24
    else:
        avg_def_km = np.mean([c.defect_count / (c.length_m/1000) for c in coils if c.length_m]) or 0.5
        defects_km = [round(avg_def_km * (1 + 0.02 * i), 4) for i in range(24)]
    forecast = [{"hour": i+1, "defects": defects_km[i]} for i in range(24)]
    return forecast

@app.get("/api/a4/fatigue_predict")
def get_a4_fatigue_predict(current_user: User = Depends(get_current_user)):
    """Global 12‑hour fatigue prediction (for UA and line admins)"""
    pred = [{"hour": i, "fatigue": round(5 + i * 0.2, 1)} for i in range(1, 13)]
    return pred

@app.get("/api/a4/fatigue_predict/inspector")
def get_fatigue_for_inspector(inspector_id: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    """Return fatigue prediction for a specific inspector."""
    if current_user.role != 'ua' and current_user.inspector_id != inspector_id:
        raise HTTPException(status_code=403, detail="Not authorized")
    inspections = db.query(Inspection).filter(Inspection.inspector_id == inspector_id).order_by(Inspection.inspection_start).all()
    if not inspections:
        pred = [{"hour": i+1, "fatigue": round(5 + i * 0.2, 1)} for i in range(12)]
    else:
        last_fatigue = inspections[-1].fatigue_score_post if inspections[-1].fatigue_score_post else 5
        pred = [{"hour": i+1, "fatigue": round(min(10, last_fatigue + i * 0.1), 1)} for i in range(12)]
    return pred

# ---------- Root endpoint ----------
@app.get("/")
def root():
    return {"message": "TSK Backend is alive. Use /docs for API documentation."}