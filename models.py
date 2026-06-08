from sqlalchemy import Column, Integer, String, Numeric, TIMESTAMP, Boolean, ForeignKey, Text, JSON
from sqlalchemy.sql import func
from database import Base

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    role = Column(String)
    line = Column(String, nullable=True)
    inspector_id = Column(String, nullable=True)
    name = Column(String)
    created_at = Column(TIMESTAMP, server_default=func.now())

class Inspector(Base):
    __tablename__ = "inspectors"
    id = Column(Integer, primary_key=True, index=True)
    inspector_id = Column(String, unique=True, index=True)
    name = Column(String)
    certified_lines = Column(JSON)
    shift_preference = Column(String)

class Coil(Base):
    __tablename__ = "coils"
    id = Column(Integer, primary_key=True, index=True)
    coil_id = Column(String)
    line = Column(String)
    start_datetime = Column(TIMESTAMP)
    end_datetime = Column(TIMESTAMP)
    length_m = Column(Numeric)
    speed_mps = Column(Numeric)
    defect_count = Column(Integer)
    defect_positions_m = Column(Text)
    created_by = Column(Integer, ForeignKey("users.id"))
    created_at = Column(TIMESTAMP, server_default=func.now())

class Inspection(Base):
    __tablename__ = "inspections"
    id = Column(Integer, primary_key=True, index=True)
    coil_id = Column(Integer, ForeignKey("coils.id"))
    inspector_id = Column(String, ForeignKey("inspectors.inspector_id"))
    inspection_start = Column(TIMESTAMP)
    inspection_end = Column(TIMESTAMP)
    fatigue_score_post = Column(Integer)
    missed_defects_estimated = Column(Integer)
    created_by = Column(Integer, ForeignKey("users.id"))
    created_at = Column(TIMESTAMP, server_default=func.now())

class AlgorithmResult(Base):
    __tablename__ = "algorithm_results"
    id = Column(Integer, primary_key=True, index=True)
    algorithm_name = Column(String)
    result_json = Column(JSON)
    computed_at = Column(TIMESTAMP, server_default=func.now())

class Schedule(Base):
    __tablename__ = "schedules"
    id = Column(Integer, primary_key=True, index=True)
    inspector_id = Column(String, ForeignKey("inspectors.inspector_id"))
    date = Column(TIMESTAMP)
    shift = Column(String)
    line = Column(String)
    rotation_flag = Column(Boolean, default=False)
    created_at = Column(TIMESTAMP, server_default=func.now())

class AuditLog(Base):
    __tablename__ = "audit_log"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"))
    action = Column(String)
    table_name = Column(String)
    record_id = Column(Integer)
    timestamp = Column(TIMESTAMP, server_default=func.now())