import os
import tempfile
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
    inspections = db.query(Inspection).offset(skip).limit(limit).all()
    return inspections

# ---------- Export endpoint (with AI sheets built‑in) ----------
@app.get("/export")
def export_results(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    coils = db.query(Coil).order_by(Coil.created_at.desc()).all()
    inspections = db.query(Inspection).all()
    inspectors = db.query(Inspector).all()

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

    # AI Sheets A1-A10 (dummy data)
    a1 = wb.create_sheet("A1 Demand Forecast")
    a1.append(["Hour", "Forecasted Defects"])
    for i in range(1, 25):
        a1.append([i, round(i * 0.5, 2)])

    a2 = wb.create_sheet("A2 Anomaly Detection")
    a2.append(["Coil ID", "Anomaly Score"])
    a2.append(["CGL-001", 0.12])
    a2.append(["CGL-002", 0.95])
    a2.append(["CAL-001", 0.03])
    a2.append(["RCL-001", 0.45])

    a3 = wb.create_sheet("A3 RL Policy")
    a3.append(["Fatigue Level", "Time on line (min)", "Recommended Action"])
    a3.append(["Low (1-3)", "0-10", "Continue"])
    a3.append(["Medium (4-6)", "10-20", "Continue"])
    a3.append(["High (7-8)", "20-30", "Rotate"])
    a3.append(["Critical (9-10)", "30+", "Rotate Immediately"])

    a4 = wb.create_sheet("A4 Fatigue Predict")
    a4.append(["Hour", "Predicted Fatigue (1-10)"])
    for i in range(1, 13):
        a4.append([i, round(5 + i * 0.2, 1)])

    a5 = wb.create_sheet("A5 DP Scheduling")
    a5.append(["Shift", "Inspectors Required"])
    a5.append(["Morning", 4])
    a5.append(["Afternoon", 3])
    a5.append(["Night", 2])

    a6 = wb.create_sheet("A6 Genetic Algorithm")
    a6.append(["Line", "Assigned Inspectors", "Optimal?"])
    a6.append(["CGL", 2, "Yes"])
    a6.append(["CAL", 2, "Yes"])
    a6.append(["RCL", 2, "Yes"])

    a7 = wb.create_sheet("A7 CUSUM Control")
    a7.append(["Sample", "CUSUM Statistic"])
    for i in range(1, 25):
        a7.append([i, round(i * 0.1, 2)])

    a8 = wb.create_sheet("A8 Monte Carlo")
    a8.append(["Risk Category", "Probability"])
    a8.append(["Low Risk", 0.70])
    a8.append(["Medium Risk", 0.20])
    a8.append(["High Risk", 0.10])

    a9 = wb.create_sheet("A9 Markov Chain")
    a9.append(["Inspector State", "Steady-State Probability (%)"])
    a9.append(["Active", 55])
    a9.append(["Fatigued", 25])
    a9.append(["Rotating", 15])
    a9.append(["Absent", 3])
    a9.append(["Training", 2])

    a10 = wb.create_sheet("A10 Live Dashboard")
    a10.append(["Line", "OEE (%)", "Utilization (%)", "Alerts"])
    a10.append(["CGL", 87.5, 92.0, "OK"])
    a10.append(["CAL", 91.2, 88.5, "OK"])
    a10.append(["RCL", 79.3, 85.0, "Fatigue Alert"])

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    return FileResponse(
        path=tmp_path,
        filename=f"TSK_Final_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- AI Results JSON Endpoints (for Flutter app) ----------
@app.get("/api/a1/forecast")
def get_a1_forecast(current_user: User = Depends(get_current_user)):
    """Return A1 demand forecast (24h) as JSON"""
    forecast = [{"hour": i, "defects": round(i * 0.5, 2)} for i in range(1, 25)]
    return forecast

@app.get("/api/a4/fatigue_predict")
def get_a4_fatigue_predict(current_user: User = Depends(get_current_user)):
    """Return 12‑hour fatigue prediction as JSON"""
    pred = [{"hour": i, "fatigue": round(5 + i * 0.2, 1)} for i in range(1, 13)]
    return pred

# ---------- Root endpoint ----------
@app.get("/")
def root():
    return {"message": "TSK Backend is alive. Use /docs for API documentation."}