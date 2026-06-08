from pydantic import BaseModel, EmailStr
from typing import Optional, List
from datetime import datetime

class UserCreate(BaseModel):
    email: EmailStr
    password: str
    role: str
    line: Optional[str] = None
    inspector_id: Optional[str] = None
    name: Optional[str] = None

class UserOut(BaseModel):
    id: int
    email: str
    role: str
    line: Optional[str] = None
    inspector_id: Optional[str] = None
    name: Optional[str] = None

class Token(BaseModel):
    access_token: str
    token_type: str

class TokenData(BaseModel):
    email: str

class CoilCreate(BaseModel):
    coil_id: str
    line: str
    start_datetime: datetime
    end_datetime: datetime
    length_m: float
    speed_mps: float
    defect_count: int
    defect_positions_m: str

class CoilOut(CoilCreate):
    id: int
    created_at: datetime

class InspectionCreate(BaseModel):
    coil_id: int
    inspector_id: str
    inspection_start: datetime
    inspection_end: datetime
    fatigue_score_post: int
    missed_defects_estimated: int

class InspectionOut(InspectionCreate):
    id: int
    created_at: datetime

class InspectorCreate(BaseModel):
    inspector_id: str
    name: str
    certified_lines: List[str]
    shift_preference: str

class InspectorOut(InspectorCreate):
    id: int