# -*- coding: utf-8 -*-
"""
models.py
Capa de datos y modelos SQLAlchemy para persistencia multiusuario y proyectos.
Compatible de forma nativa con SQLite (local) y PostgreSQL (producción / nube).
"""

import os
import json
from datetime import datetime
from sqlalchemy import (
    create_engine, Column, Integer, String, Text, DateTime,
    ForeignKey, Boolean, inspect, text
)
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


class Group(Base):
    __tablename__ = "groups"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(120), unique=True, nullable=False, index=True)
    description = Column(String(255), nullable=True, default="")
    created_at = Column(DateTime, default=datetime.utcnow)

    users = relationship("User", back_populates="group")

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description or "",
            "member_count": len(self.users) if self.users else 0,
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M") if self.created_at else ""
        }


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(80), unique=True, nullable=False, index=True)
    email = Column(String(120), unique=True, nullable=True)
    password_hash = Column(String(256), nullable=False)
    role = Column(String(20), default="user", nullable=False)  # "admin" ou "user"
    is_approved = Column(Boolean, default=False, nullable=False)
    group_id = Column(Integer, ForeignKey("groups.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    group = relationship("Group", back_populates="users")
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
            "role": self.role or "user",
            "is_admin": (self.role == "admin"),
            "is_approved": bool(self.is_approved),
            "group_id": self.group_id,
            "group_name": self.group.name if self.group else None,
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

    # Datos técnicos específicos de Habitabilidade de Vivendas (NHV - Decreto 29/2010)
    nhv_data = Column(Text, nullable=False, default="{}")

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

    def get_nhv_dict(self):
        try:
            return json.loads(self.nhv_data) if self.nhv_data else {}
        except Exception:
            return {}

    def set_nhv_dict(self, d):
        self.nhv_data = json.dumps(d, ensure_ascii=False)

    def to_dict(self):
        common = self.get_common_dict()
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description or "",
            "user_id": self.user_id,
            "owner_username": self.user.username if self.user else "Descoñecido",
            "group_id": self.user.group_id if (self.user and self.user.group_id) else None,
            "group_name": self.user.group.name if (self.user and self.user.group) else None,
            "common_data": common,
            "ebss_data": self.get_ebss_dict(),
            "residuos_data": self.get_residuos_dict(),
            "nhv_data": self.get_nhv_dict(),
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M") if self.updated_at else "",
            "poboacion": common.get("poboacion", ""),
            "tipo_obra": common.get("tipo_obra", "")
        }


def init_db():
    Base.metadata.create_all(bind=engine)

    # Migración segura e idempotente para columnas novas
    try:
        inspector = inspect(engine)
        if "users" in inspector.get_table_names():
            existing_columns = [col["name"] for col in inspector.get_columns("users")]
            is_pg = engine.dialect.name == "postgresql"
            bool_col_type = "BOOLEAN DEFAULT FALSE" if is_pg else "BOOLEAN DEFAULT 0"

            with engine.begin() as conn:
                if "role" not in existing_columns:
                    conn.execute(text("ALTER TABLE users ADD COLUMN role VARCHAR(20) DEFAULT 'user'"))
                if "is_approved" not in existing_columns:
                    conn.execute(text(f"ALTER TABLE users ADD COLUMN is_approved {bool_col_type}"))
                if "group_id" not in existing_columns:
                    conn.execute(text("ALTER TABLE users ADD COLUMN group_id INTEGER REFERENCES groups(id)"))

                # Garantir que o usuario 'pgarat' sexa sempre Administrador e estea aprobado
                conn.execute(
                    text("UPDATE users SET role = 'admin', is_approved = :appr WHERE LOWER(username) = 'pgarat'"),
                    {"appr": True}
                )

                # Asegurar que usuarios preexistentes queden aprobados para non bloquear instalacións en marcha
                conn.execute(
                    text("UPDATE users SET is_approved = :appr WHERE is_approved IS NULL"),
                    {"appr": True}
                )

        if "projects" in inspector.get_table_names():
            proj_cols = [col["name"] for col in inspector.get_columns("projects")]
            with engine.begin() as conn:
                if "nhv_data" not in proj_cols:
                    conn.execute(text("ALTER TABLE projects ADD COLUMN nhv_data TEXT DEFAULT '{}'"))
    except Exception as e:
        print(f"[AVISO] Erro durante a comprobación/migración de columnas en init_db: {e}", flush=True)


def get_db():
    db = SessionLocal()
    try:
        return db
    finally:
        pass
