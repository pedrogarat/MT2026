# -*- coding: utf-8 -*-
"""
models.py
Capa de datos y modelos SQLAlchemy para persistencia multiusuario y proyectos.
Compatible de forma nativa con SQLite (local) y PostgreSQL (producción / nube).
"""

import os
import json
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker, relationship, scoped_session
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
raw_db_url = os.environ.get("DATABASE_URL", f"sqlite:///{os.path.join(BASE_DIR, 'asistente_ebss.db')}")

# Compatibilidad con URLs de PostgreSQL en plataformas como Render/Heroku (postgres:// -> postgresql://)
if raw_db_url.startswith("postgres://"):
    raw_db_url = raw_db_url.replace("postgres://", "postgresql://", 1)

engine = create_engine(
    raw_db_url,
    connect_args={"check_same_thread": False} if raw_db_url.startswith("sqlite") else {}
)

SessionLocal = scoped_session(sessionmaker(autocommit=False, autoflush=False, bind=engine))
Base = declarative_base()


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(80), unique=True, nullable=False, index=True)
    email = Column(String(120), unique=True, nullable=True)
    password_hash = Column(String(256), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    projects = relationship("Project", back_populates="user", cascade="all, delete-orphan", order_by="desc(Project.updated_at)")

    def set_password(self, password: str):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "email": self.email,
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M") if self.created_at else ""
        }


class Project(Base):
    __tablename__ = "projects"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False, index=True)
    name = Column(String(200), nullable=False)
    description = Column(String(500), nullable=True, default="")
    
    # Datos comunes del proyecto (Objeto, Situación, Concello, Promotor, Arquitecto, PEM, Duración, Superficies)
    common_data = Column(Text, nullable=False, default="{}")
    
    # Datos técnicos específicos de EBSS (Condicionantes, Fases, Instalaciones, Asistencia, Maquinaria)
    ebss_data = Column(Text, nullable=False, default="{}")
    
    # Datos técnicos específicos de Gestión de Residuos (Selección de RCDs, Firmas)
    residuos_data = Column(Text, nullable=False, default="{}")

    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    user = relationship("User", back_populates="projects")

    def get_common_dict(self):
        try:
            return json.loads(self.common_data) if self.common_data else {}
        except Exception:
            return {}

    def set_common_dict(self, d):
        self.common_data = json.dumps(d, ensure_ascii=False)

    def get_ebss_dict(self):
        try:
            return json.loads(self.ebss_data) if self.ebss_data else {}
        except Exception:
            return {}

    def set_ebss_dict(self, d):
        self.ebss_data = json.dumps(d, ensure_ascii=False)

    def get_residuos_dict(self):
        try:
            return json.loads(self.residuos_data) if self.residuos_data else {}
        except Exception:
            return {}

    def set_residuos_dict(self, d):
        self.residuos_data = json.dumps(d, ensure_ascii=False)

    def to_dict(self):
        common = self.get_common_dict()
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description or "",
            "common_data": common,
            "ebss_data": self.get_ebss_dict(),
            "residuos_data": self.get_residuos_dict(),
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M") if self.updated_at else "",
            "poboacion": common.get("poboacion", ""),
            "tipo_obra": common.get("tipo_obra", "")
        }


def init_db():
    Base.metadata.create_all(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        return db
    finally:
        pass
