import base64
import hashlib
import hmac
import os
import secrets
from datetime import datetime, time
from zoneinfo import ZoneInfo

from sqlalchemy import Boolean, Column, DateTime, ForeignKey, Integer, String, Table, Time, create_engine, inspect, select, text
from sqlalchemy.orm import DeclarativeBase, Session, joinedload, relationship, sessionmaker

DATABASE_URL = os.getenv("DATABASE_URL", "postgresql+psycopg://purchasing:purchasing@postgres:5432/purchasing")
DEFAULT_TIMEZONE = os.getenv("APP_TIMEZONE", "Asia/Jakarta")


class Base(DeclarativeBase):
    pass


user_roles = Table("user_roles", Base.metadata, Column("user_id", ForeignKey("users.id", ondelete="CASCADE"), primary_key=True), Column("role_id", ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True))
role_profile_roles = Table("role_profile_roles", Base.metadata, Column("profile_id", ForeignKey("role_profiles.id", ondelete="CASCADE"), primary_key=True), Column("role_id", ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True))
role_permissions = Table("role_permissions", Base.metadata, Column("role_id", ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True), Column("permission_id", ForeignKey("permissions.id", ondelete="CASCADE"), primary_key=True))
page_roles = Table("page_roles", Base.metadata, Column("page_id", ForeignKey("pages.id", ondelete="CASCADE"), primary_key=True), Column("role_id", ForeignKey("roles.id", ondelete="CASCADE"), primary_key=True))


class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    username = Column(String(120), unique=True, nullable=False, index=True)
    full_name = Column(String(120), nullable=False)
    password_hash = Column(String(255), nullable=False)
    is_active = Column(Boolean, nullable=False, default=True)
    timezone = Column(String(64), nullable=False, default=DEFAULT_TIMEZONE)
    login_after = Column(Time, nullable=True)
    login_before = Column(Time, nullable=True)
    last_login_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    roles = relationship("Role", secondary=user_roles, back_populates="users")


class Role(Base):
    __tablename__ = "roles"
    id = Column(Integer, primary_key=True)
    name = Column(String(64), unique=True, nullable=False)
    description = Column(String(255), nullable=False, default="")
    is_active = Column(Boolean, nullable=False, default=True)
    users = relationship("User", secondary=user_roles, back_populates="roles")
    permissions = relationship("Permission", secondary=role_permissions, back_populates="roles")
    profiles = relationship("RoleProfile", secondary=role_profile_roles, back_populates="roles")


class RoleProfile(Base):
    __tablename__ = "role_profiles"
    id = Column(Integer, primary_key=True)
    name = Column(String(80), unique=True, nullable=False)
    description = Column(String(255), nullable=False, default="")
    roles = relationship("Role", secondary=role_profile_roles, back_populates="profiles")


class Page(Base):
    __tablename__ = "pages"
    id = Column(Integer, primary_key=True)
    path = Column(String(120), unique=True, nullable=False)
    label = Column(String(120), nullable=False)
    kind = Column(String(20), nullable=False, default="Page")
    roles = relationship("Role", secondary=page_roles)


class Permission(Base):
    __tablename__ = "permissions"
    id = Column(Integer, primary_key=True)
    code = Column(String(64), unique=True, nullable=False)
    label = Column(String(120), nullable=False)
    roles = relationship("Role", secondary=role_permissions, back_populates="permissions")


engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
PERMISSIONS = {"dashboard.read": "Dashboard", "upload.create": "Upload Template", "audit.read": "Audit Admin", "settings.write": "Settings", "users.manage": "User Management", "report.export": "Export Report", "restock.read": "Restock"}
PAGES = {"/dashboard": ("Dashboard", "Page"), "/": ("Upload Template", "Page"), "/audit": ("Audit Admin", "Page"), "/settings": ("Settings", "Page"), "/users": ("User Management", "Page"), "/restock": ("Restock", "Page"), "/export": ("Export Report", "Report")}


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode(), salt=salt, n=16384, r=8, p=1)
    return f"scrypt${base64.b64encode(salt).decode()}${base64.b64encode(digest).decode()}"


def verify_password(password: str, password_hash: str) -> bool:
    try:
        algorithm, salt_text, digest_text = password_hash.split("$", 2)
        return algorithm == "scrypt" and hmac.compare_digest(hashlib.scrypt(password.encode(), salt=base64.b64decode(salt_text), n=16384, r=8, p=1), base64.b64decode(digest_text))
    except (ValueError, TypeError):
        return False


def initialize_auth() -> None:
    inspector = inspect(engine)
    with engine.begin() as connection:
        if inspector.has_table("users"):
            columns = {column["name"] for column in inspector.get_columns("users")}
            for name, sql in {"timezone": "ALTER TABLE users ADD COLUMN timezone VARCHAR(64) NOT NULL DEFAULT 'Asia/Jakarta'", "login_after": "ALTER TABLE users ADD COLUMN login_after TIME", "login_before": "ALTER TABLE users ADD COLUMN login_before TIME", "last_login_at": "ALTER TABLE users ADD COLUMN last_login_at TIMESTAMP"}.items():
                if name not in columns:
                    connection.execute(text(sql))
        if inspector.has_table("roles"):
            role_columns = {column["name"] for column in inspector.get_columns("roles")}
            if "is_active" not in role_columns:
                connection.execute(text("ALTER TABLE roles ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT TRUE"))
    Base.metadata.create_all(engine)
    with SessionLocal() as db:
        permissions = {}
        for code, label in PERMISSIONS.items():
            permission = db.scalar(select(Permission).where(Permission.code == code))
            if not permission:
                permission = Permission(code=code, label=label)
                db.add(permission)
            permissions[code] = permission
        manager = db.scalar(select(Role).where(Role.name == "System Manager"))
        if not manager:
            manager = Role(name="System Manager", description="Akses penuh aplikasi")
            db.add(manager)
        db.flush()
        manager.permissions = list(permissions.values())
        for path, (label, kind) in PAGES.items():
            page = db.scalar(select(Page).where(Page.path == path))
            if not page:
                page = Page(path=path, label=label, kind=kind)
                db.add(page)
                db.flush()
            if manager not in page.roles:
                page.roles.append(manager)
        username = os.getenv("ADMIN_USERNAME", "admin").strip().lower()
        password = os.getenv("ADMIN_PASSWORD", "")
        if password and not db.scalar(select(User).where(User.username == username)):
            db.add(User(username=username, full_name="System Administrator", password_hash=hash_password(password), roles=[manager], timezone=DEFAULT_TIMEZONE))
        db.commit()


def get_user(db: Session, user_id: int | None) -> User | None:
    if not user_id:
        return None
    user = db.scalar(select(User).options(joinedload(User.roles).joinedload(Role.permissions)).where(User.id == user_id))
    return user if user and user.is_active else None


def has_permission(user: User | None, code: str) -> bool:
    return bool(user and any(permission.code == code for role in user.roles if role.is_active for permission in role.permissions))


def can_login_now(user: User) -> bool:
    if not user.login_after or not user.login_before:
        return True
    try:
        now = datetime.now(ZoneInfo(user.timezone or DEFAULT_TIMEZONE)).time()
    except Exception:
        now = datetime.now(ZoneInfo(DEFAULT_TIMEZONE)).time()
    if user.login_after <= user.login_before:
        return user.login_after <= now <= user.login_before
    return now >= user.login_after or now <= user.login_before


def parse_time(value: str | None) -> time | None:
    return datetime.strptime(value, "%H:%M").time() if value else None
