import os
import tempfile
import math
from datetime import datetime
from typing import List

from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.responses import FileResponse
from jose import jwt, JWTError
from sqlalchemy.orm import Session
from dotenv import load_dotenv

from database import get_db, engine
from models import Base, User, Coil, Inspection, Inspector
from schemas import UserCreate, UserOut, Token, CoilCreate, CoilOut, InspectionCreate, InspectionOut
from auth import get_password_hash, verify_password, create_access_token, SECRET_KEY, ALGORITHM

# openpyxl for Excel export
from openpyxl import Workbook

# Import AI engine
from tsk_final_engine import add_ai_sheets_to_workbook
import pandas as pd
import numpy as np

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

# Helper: get current user
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

# ---------- Auth ----------
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
    inspections = db.query(Inspection).offset(skip).limit(limit).all()
    return inspections

# ---------- Helper to compute line statistics from coils ----------
def compute_line_stats(coils, inspectors):
    line_stats = {}
    for line in ["CGL", "CAL", "RCL"]:
        line_coils = [c for c in coils if c.line == line]
        n_coils = len(line_coils)
        if n_coils == 0:
            # Default values if no coils exist
            wl_avg = 0.76
            defects_km_avg = 0.5
            cv = 0.3
            fat_avg = 6.5
        else:
            # Calculate average W_l (s/m) from length and speed
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
            cv = np.std(defects_km)/defects_km_avg if defects_km_avg>0 and len(defects_km)>1 else 0.3
            # Average fatigue from inspections (not easily available here; use placeholder)
            fat_avg = 6.5
        speed = {"CGL":150, "CAL":200, "RCL":250}[line]
        raw = (speed * wl_avg) / 60  # minutes per coil approx
        n_base = max(1, math.ceil(raw))
        n_peak = math.ceil(n_base * 1.3)
        line_stats[line] = {
            "wl_avg": round(wl_avg,4),
            "defects_km_avg": round(defects_km_avg,4),
            "n_coils": n_coils,
            "speed": speed,
            "n_base": n_base,
            "n_peak": n_peak,
            "cv": round(cv,4),
            "fat_avg": round(fat_avg,2)
        }
    return line_stats

# ---------- Export ----------
@app.get("/export")
def export_results(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    coils = db.query(Coil).order_by(Coil.created_at.desc()).all()
    inspections = db.query(Inspection).all()
    inspectors = db.query(Inspector).all()

    # Convert to DataFrames for AI engine
    coils_df = pd.DataFrame([{k: getattr(c, k) for k in ['id','coil_id','line','start_datetime','end_datetime',
                                                          'length_m','speed_mps','defect_count','defect_positions_m']} for c in coils])
    inspections_df = pd.DataFrame([{k: getattr(i, k) for k in ['id','coil_id','inspector_id','inspection_start',
                                                                'inspection_end','fatigue_score_post','missed_defects_estimated']} for i in inspections])
    inspectors_df = pd.DataFrame([{k: getattr(ins, k) for k in ['inspector_id','name','certified_lines','shift_preference']} for ins in inspectors])

    # Compute line statistics for AI engine
    line_stats = compute_line_stats(coils, inspectors)

    # Create base workbook with standard sheets
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.append(["TSK COIL INSPECTION - DASHBOARD"])
    ws.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
    ws.append([])
    ws.append(["Total Coils", len(coils)])
    ws.append(["Total Inspections", len(inspections)])
    ws.append(["Inspectors", len(inspectors)])

    ws2 = wb.create_sheet("Coil Detail")
    ws2.append(["ID", "Coil ID", "Line", "Start Time", "End Time", "Length (m)", "Speed (m/s)", "Defect Count", "Defect Positions"])
    for c in coils:
        ws2.append([c.id, c.coil_id, c.line, c.start_datetime, c.end_datetime, c.length_m, c.speed_mps, c.defect_count, c.defect_positions_m])

    ws3 = wb.create_sheet("Inspection Log")
    ws3.append(["Inspection ID", "Coil ID", "Inspector ID", "Start", "End", "Fatigue Score", "Missed Defects"])
    for insp in inspections:
        coil = db.query(Coil).filter(Coil.id == insp.coil_id).first()
        coil_id_str = coil.coil_id if coil else "Unknown"
        ws3.append([insp.id, coil_id_str, insp.inspector_id, insp.inspection_start, insp.inspection_end, insp.fatigue_score_post, insp.missed_defects_estimated])

    ws4 = wb.create_sheet("Inspector Matrix")
    ws4.append(["Inspector ID", "Name", "Certified Lines", "Shift Preference"])
    for ins in inspectors:
        ws4.append([ins.inspector_id, ins.name, ", ".join(ins.certified_lines) if ins.certified_lines else "", ins.shift_preference])

    # Add AI sheets A1-A10 using the engine (which appends to the same workbook)
    try:
        wb = add_ai_sheets_to_workbook(wb, coils_df, inspections_df, inspectors_df, line_stats, avail_insp=len(inspectors))
    except Exception as e:
        # If AI sheets fail, still return the basic workbook
        print(f"AI sheets error: {e}")

    # Save to a temporary file and return
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    return FileResponse(
        path=tmp_path,
        filename=f"TSK_Final_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/")
def root():
    return {"message": "TSK Backend is alive. Use /docs for API documentation."}