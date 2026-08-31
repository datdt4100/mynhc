import os
import re
import json
import io
import secrets
import string
import unicodedata
from functools import wraps
from datetime import datetime

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, jsonify, send_file, flash, Response
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import (
    create_engine, MetaData, Table, Column, Integer, Text,
    insert, select, update, delete, and_, or_, func, text
)
import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///classreg.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, future=True)
metadata = MetaData()

# ---------------------------------------------------------------------------
# Table definitions
# ---------------------------------------------------------------------------

teachers = Table(
    "teachers", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("full_name", Text, nullable=False),
    Column("cccd", Text, unique=True, nullable=False),
    Column("gender", Text, nullable=False),           # Nam / Nữ
    Column("subject_group", Text, nullable=False),    # Tổ bộ môn
    Column("email", Text, nullable=True),
    Column("password_hash", Text, nullable=True),
    Column("is_first_login", Integer, default=1),
    Column("must_change_password", Integer, default=0),
)

students = Table(
    "students", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("full_name", Text, nullable=False),
    Column("cccd", Text, unique=True, nullable=False),
    Column("class_name", Text, nullable=False),
    Column("grade", Integer, nullable=False),         # 10 / 11 / 12
    Column("email", Text, nullable=True),
    Column("password_hash", Text, nullable=True),
    Column("is_first_login", Integer, default=1),
    Column("must_change_password", Integer, default=0),
)

classes = Table(
    "classes", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("teacher_id", Integer, nullable=False),
    Column("grade", Integer, nullable=False),
    Column("duration", Integer, nullable=False),      # 1-4 tiết
    Column("day_of_week", Integer, nullable=False),   # 2-7
    Column("session_type", Text, nullable=False),     # morning / afternoon
    Column("start_session", Integer, nullable=False), # 1-4
    Column("subject", Text, nullable=True),
    Column("location", Text, nullable=True),
    Column("max_capacity", Integer, nullable=True),
    Column("extra_data", Text, nullable=True),        # JSON blob
    Column("is_published", Integer, default=0),
    Column("created_at", Text, default="CURRENT_TIMESTAMP"),
)

enrollments = Table(
    "enrollments", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("student_id", Integer, nullable=False),
    Column("class_id", Integer, nullable=False),
    Column("enrolled_at", Text, default="CURRENT_TIMESTAMP"),
)

settings_table = Table(
    "settings", metadata,
    Column("key", Text, primary_key=True),
    Column("value", Text, nullable=True),
)

# ---------------------------------------------------------------------------
# DB init
# ---------------------------------------------------------------------------

def init_db():
    metadata.create_all(engine)
    with engine.connect() as conn:
        # Migrate: add columns if not exists (for existing DBs)
        for stmt in [
            "ALTER TABLE teachers ADD COLUMN must_change_password INTEGER DEFAULT 0",
            "ALTER TABLE students ADD COLUMN must_change_password INTEGER DEFAULT 0",
            "ALTER TABLE students ADD COLUMN email TEXT",
        ]:
            try:
                conn.execute(text(stmt))
                conn.commit()
            except Exception:
                conn.rollback()

        # Seed default settings
        for key, value in [("admin_password", "admin123"),
                           ("teacher_reg_open", "0"),
                           ("student_reg_open", "0")]:
            try:
                conn.execute(
                    insert(settings_table).values(key=key, value=value)
                )
                conn.commit()
            except Exception:
                conn.rollback()

with app.app_context():
    init_db()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Allowed special chars (excludes ' " ` ; \ which are DB-dangerous)
_PW_ALLOWED_SPECIALS = r"!@#$%^&*()\-_+=\[\]{}|<>,.?/~"
_PW_PATTERN = re.compile(r'^[A-Za-z0-9' + _PW_ALLOWED_SPECIALS + r']+$')


def validate_password(pw: str):
    """Normalize and validate password. Returns (ok, error_msg)."""
    pw = unicodedata.normalize('NFKC', pw).strip()
    if len(pw) < 8:
        return False, "Mật khẩu phải có ít nhất 8 ký tự."
    if not _PW_PATTERN.match(pw):
        return False, (
            "Mật khẩu chỉ được chứa chữ hoa, chữ thường, chữ số "
            "và ký tự đặc biệt hợp lệ (!@#$%^&*()-_+=[]{}|<>,.?/~). "
            "Không dùng dấu nháy (', \"), dấu chấm phẩy (;) hoặc dấu gạch chéo (\\)."
        )
    return True, None


def normalize_password(pw: str) -> str:
    return unicodedata.normalize('NFKC', pw).strip()


def generate_temp_password(length: int = 12) -> str:
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def get_setting(key, default=None):
    with engine.connect() as conn:
        row = conn.execute(
            select(settings_table.c.value).where(settings_table.c.key == key)
        ).fetchone()
    return row[0] if row else default


def set_setting(key, value):
    with engine.connect() as conn:
        existing = conn.execute(
            select(settings_table.c.key).where(settings_table.c.key == key)
        ).fetchone()
        if existing:
            conn.execute(
                update(settings_table).where(settings_table.c.key == key).values(value=value)
            )
        else:
            conn.execute(insert(settings_table).values(key=key, value=value))
        conn.commit()


def _is_admin():
    return session.get("is_admin") is True


def day_name(n):
    mapping = {2: "Thứ 2", 3: "Thứ 3", 4: "Thứ 4",
               5: "Thứ 5", 6: "Thứ 6", 7: "Thứ 7"}
    return mapping.get(int(n), str(n))


def session_label(session_type, start, duration):
    label = "Sáng" if session_type == "morning" else "Chiều"
    end = int(start) + int(duration) - 1
    if int(duration) == 1:
        return f"{label}, tiết {start}"
    return f"{label}, tiết {start}-{end}"


def _time_conflict(student_id, new_class):
    """Return the conflicting class row if any, else None."""
    with engine.connect() as conn:
        enrolled = conn.execute(
            select(classes).join(
                enrollments, classes.c.id == enrollments.c.class_id
            ).where(enrollments.c.student_id == student_id)
        ).fetchall()

    new_day = int(new_class["day_of_week"])
    new_session = new_class["session_type"]
    new_start = int(new_class["start_session"])
    new_end = new_start + int(new_class["duration"]) - 1

    for c in enrolled:
        if int(c.day_of_week) != new_day:
            continue
        if c.session_type != new_session:
            continue
        c_start = int(c.start_session)
        c_end = c_start + int(c.duration) - 1
        # Overlap check
        if new_start <= c_end and c_start <= new_end:
            return c
    return None


def teacher_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get("user_type") != "teacher":
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated


def student_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get("user_type") != "student":
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not _is_admin():
            return redirect(url_for("admin_login_page"))
        return f(*args, **kwargs)
    return decorated

# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return redirect(url_for("login_page"))


@app.route("/login")
def login_page():
    return render_template("login.html")


@app.route("/login/step1", methods=["POST"])
def login_step1():
    data = request.get_json(force=True)
    full_name = (data.get("full_name") or "").strip()
    cccd = (data.get("cccd") or "").strip()

    if not full_name or not cccd:
        return jsonify(ok=False, error="Vui lòng nhập đầy đủ thông tin.")

    with engine.connect() as conn:
        teacher = conn.execute(
            select(teachers).where(
                and_(teachers.c.full_name == full_name, teachers.c.cccd == cccd)
            )
        ).fetchone()
        if teacher:
            title = "Cô" if teacher.gender == "Nữ" else "Thầy"
            return jsonify(
                ok=True,
                user_type="teacher",
                full_name=teacher.full_name,
                gender=teacher.gender,
                title=title,
                is_first_login=teacher.is_first_login,
            )

        student = conn.execute(
            select(students).where(
                and_(students.c.full_name == full_name, students.c.cccd == cccd)
            )
        ).fetchone()
        if student:
            return jsonify(
                ok=True,
                user_type="student",
                full_name=student.full_name,
                gender=None,
                title="Học sinh",
                is_first_login=student.is_first_login,
            )

    return jsonify(ok=False, error="Không tìm thấy tài khoản phù hợp.")


@app.route("/login/step2", methods=["POST"])
def login_step2():
    data = request.get_json(force=True)
    full_name = (data.get("full_name") or "").strip()
    cccd = (data.get("cccd") or "").strip()
    user_type = data.get("user_type")
    password = normalize_password(data.get("password") or "")
    email = (data.get("email") or "").strip()

    if user_type == "teacher":
        with engine.connect() as conn:
            teacher = conn.execute(
                select(teachers).where(
                    and_(teachers.c.full_name == full_name, teachers.c.cccd == cccd)
                )
            ).fetchone()

        if not teacher:
            return jsonify(ok=False, error="Không tìm thấy giáo viên.")

        if teacher.is_first_login:
            # First login: set email + password
            if not email:
                return jsonify(ok=False, error="Vui lòng nhập email.")
            ok, err = validate_password(password)
            if not ok:
                return jsonify(ok=False, error=err)
            pw_hash = generate_password_hash(password)
            with engine.connect() as conn:
                conn.execute(
                    update(teachers).where(teachers.c.id == teacher.id).values(
                        email=email,
                        password_hash=pw_hash,
                        is_first_login=0,
                        must_change_password=0,
                    )
                )
                conn.commit()
        else:
            if not teacher.password_hash:
                return jsonify(ok=False, error="Tài khoản chưa được thiết lập mật khẩu.")
            if not check_password_hash(teacher.password_hash, password):
                return jsonify(ok=False, error="Sai mật khẩu.")

        session.clear()
        session["user_type"] = "teacher"
        session["user_id"] = teacher.id
        session["full_name"] = teacher.full_name
        session["gender"] = teacher.gender
        session["subject_group"] = teacher.subject_group

        must_change = getattr(teacher, 'must_change_password', 0) or 0
        if must_change and not teacher.is_first_login:
            return jsonify(ok=True, must_change=True, redirect=url_for("change_password_page"))
        return jsonify(ok=True, redirect=url_for("teacher_dashboard"))

    elif user_type == "student":
        with engine.connect() as conn:
            student = conn.execute(
                select(students).where(
                    and_(students.c.full_name == full_name, students.c.cccd == cccd)
                )
            ).fetchone()

        if not student:
            return jsonify(ok=False, error="Không tìm thấy học sinh.")

        if student.is_first_login:
            # First login: set password (no email required for students)
            ok, err = validate_password(password)
            if not ok:
                return jsonify(ok=False, error=err)
            pw_hash = generate_password_hash(password)
            with engine.begin() as conn:
                conn.execute(
                    update(students).where(students.c.id == student.id).values(
                        password_hash=pw_hash,
                        is_first_login=0,
                        must_change_password=0,
                    )
                )
        else:
            if not student.password_hash:
                return jsonify(ok=False, error="Tài khoản chưa được thiết lập mật khẩu.")
            if not check_password_hash(student.password_hash, password):
                return jsonify(ok=False, error="Sai mật khẩu.")

        session.clear()
        session["user_type"] = "student"
        session["user_id"] = student.id
        session["full_name"] = student.full_name
        session["grade"] = student.grade

        must_change = getattr(student, 'must_change_password', 0) or 0
        if must_change and not student.is_first_login:
            return jsonify(ok=True, must_change=True, redirect=url_for("change_password_page"))
        return jsonify(ok=True, redirect=url_for("student_dashboard"))

    return jsonify(ok=False, error="Loại tài khoản không hợp lệ.")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login_page"))


@app.route("/change-password", methods=["GET"])
def change_password_page():
    user_type = session.get("user_type")
    if user_type not in ("teacher", "student"):
        return redirect(url_for("login_page"))
    return render_template("change_password.html")


@app.route("/change-password", methods=["POST"])
def change_password_submit():
    user_type = session.get("user_type")
    user_id   = session.get("user_id")
    if user_type not in ("teacher", "student") or not user_id:
        return jsonify(ok=False, error="Phiên đăng nhập hết hạn.")

    data = request.get_json(force=True)
    new_pw  = normalize_password(data.get("new_password") or "")
    confirm = normalize_password(data.get("confirm_password") or "")

    ok, err = validate_password(new_pw)
    if not ok:
        return jsonify(ok=False, error=err)
    if new_pw != confirm:
        return jsonify(ok=False, error="Mật khẩu xác nhận không khớp.")

    pw_hash = generate_password_hash(new_pw)
    table = teachers if user_type == "teacher" else students
    with engine.begin() as conn:
        conn.execute(
            update(table).where(table.c.id == user_id).values(
                password_hash=pw_hash,
                must_change_password=0,
            )
        )

    redirect_url = url_for("teacher_dashboard") if user_type == "teacher" else url_for("student_dashboard")
    return jsonify(ok=True, redirect=redirect_url)

# ---------------------------------------------------------------------------
# Teacher routes
# ---------------------------------------------------------------------------

@app.route("/teacher")
@teacher_required
def teacher_dashboard():
    teacher_id = session["user_id"]
    with engine.connect() as conn:
        teacher_row = conn.execute(
            select(teachers).where(teachers.c.id == teacher_id)
        ).fetchone()

        teacher_classes = conn.execute(
            select(classes).where(classes.c.teacher_id == teacher_id)
            .order_by(classes.c.grade, classes.c.day_of_week, classes.c.session_type, classes.c.start_session)
        ).fetchall()

        enrollment_counts = {}
        for c in teacher_classes:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == c.id)
            ).scalar()
            enrollment_counts[c.id] = cnt

    # Build schedule grid for Outlook-style calendar
    schedule = {}
    for c in teacher_classes:
        k = f"{c.day_of_week}_{c.session_type}_{c.start_session}"
        schedule[k] = {"status": "start", "cls": c, "rowspan": c.duration}
        for t in range(c.start_session + 1, c.start_session + c.duration):
            schedule[f"{c.day_of_week}_{c.session_type}_{t}"] = {"status": "blocked"}

    teacher_reg_open = get_setting("teacher_reg_open", "0") == "1"
    return render_template(
        "teacher/dashboard.html",
        teacher=teacher_row,
        classes=teacher_classes,
        enrollment_counts=enrollment_counts,
        schedule=schedule,
        teacher_reg_open=teacher_reg_open,
        day_name=day_name,
        session_label=session_label,
    )


@app.route("/teacher/classes", methods=["POST"])
@teacher_required
def teacher_register_class():
    if get_setting("teacher_reg_open", "0") != "1":
        return jsonify(ok=False, error="Chưa mở đăng ký lớp.")

    data = request.get_json(force=True)
    teacher_id = session["user_id"]

    try:
        grade = int(data["grade"])
        duration = int(data["duration"])
        day_of_week = int(data["day_of_week"])
        session_type = data["session_type"]
        start_session = int(data["start_session"])
    except (KeyError, ValueError, TypeError):
        return jsonify(ok=False, error="Dữ liệu không hợp lệ.")

    if grade not in (10, 11, 12):
        return jsonify(ok=False, error="Khối phải là 10, 11 hoặc 12.")
    if not (1 <= duration <= 4):
        return jsonify(ok=False, error="Số tiết phải từ 1 đến 4.")
    if not (2 <= day_of_week <= 7):
        return jsonify(ok=False, error="Thứ phải từ 2 đến 7.")
    if session_type not in ("morning", "afternoon"):
        return jsonify(ok=False, error="Buổi không hợp lệ.")
    if not (1 <= start_session <= 4):
        return jsonify(ok=False, error="Tiết bắt đầu phải từ 1 đến 4.")
    if start_session + duration - 1 > 4:
        return jsonify(ok=False, error="Tiết kết thúc vượt quá tiết 4.")

    subject = session.get("subject_group") or (data.get("subject") or "").strip() or None

    with engine.connect() as conn:
        result = conn.execute(
            insert(classes).values(
                teacher_id=teacher_id,
                grade=grade,
                duration=duration,
                day_of_week=day_of_week,
                session_type=session_type,
                start_session=start_session,
                subject=subject,
                location=None,
                max_capacity=None,
                extra_data=None,
                is_published=1,
                created_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
        conn.commit()
        class_id = result.inserted_primary_key[0]

    return jsonify(ok=True, class_id=class_id)


@app.route("/teacher/classes/<int:class_id>", methods=["DELETE"])
@teacher_required
def teacher_delete_class(class_id):
    if get_setting("teacher_reg_open", "0") != "1":
        return jsonify(ok=False, error="Đăng ký lớp đã đóng, không thể thực hiện thay đổi.")
    teacher_id = session["user_id"]
    with engine.connect() as conn:
        cls = conn.execute(
            select(classes).where(
                and_(classes.c.id == class_id, classes.c.teacher_id == teacher_id)
            )
        ).fetchone()
        if not cls:
            return jsonify(ok=False, error="Không tìm thấy lớp.")

        enroll_count = conn.execute(
            select(func.count()).where(enrollments.c.class_id == class_id)
        ).scalar()
        if enroll_count > 0:
            return jsonify(ok=False, error="Lớp đã có học sinh đăng ký, không thể xóa.")

        conn.execute(delete(classes).where(classes.c.id == class_id))
        conn.commit()

    return jsonify(ok=True)


@app.route("/teacher/classes/<int:class_id>/students")
@teacher_required
def teacher_class_students(class_id):
    teacher_id = session["user_id"]
    with engine.connect() as conn:
        cls = conn.execute(
            select(classes).where(
                and_(classes.c.id == class_id, classes.c.teacher_id == teacher_id)
            )
        ).fetchone()
        if not cls:
            return jsonify(ok=False, error="Không tìm thấy lớp.")

        enrolled = conn.execute(
            select(students, enrollments.c.enrolled_at)
            .join(enrollments, students.c.id == enrollments.c.student_id)
            .where(enrollments.c.class_id == class_id)
            .order_by(students.c.full_name)
        ).fetchall()

    return jsonify(ok=True, students=[
        {
            "id": s.id,
            "full_name": s.full_name,
            "class_name": s.class_name,
            "grade": s.grade,
            "enrolled_at": s.enrolled_at,
        }
        for s in enrolled
    ])

# ---------------------------------------------------------------------------
# Student routes
# ---------------------------------------------------------------------------

@app.route("/student")
@student_required
def student_dashboard():
    student_id = session["user_id"]
    student_grade = session.get("grade")

    with engine.connect() as conn:
        student_row = conn.execute(
            select(students).where(students.c.id == student_id)
        ).fetchone()

        published_classes = conn.execute(
            select(classes, teachers.c.full_name.label("teacher_name"),
                   teachers.c.subject_group)
            .join(teachers, classes.c.teacher_id == teachers.c.id)
            .where(and_(
                classes.c.is_published == 1,
                classes.c.grade == student_grade,
                classes.c.location.isnot(None),
                classes.c.location != '',
                classes.c.max_capacity.isnot(None),
            ))
            .order_by(classes.c.day_of_week, classes.c.session_type, classes.c.start_session)
        ).fetchall()

        enrollment_counts = {}
        for c in published_classes:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == c.id)
            ).scalar()
            enrollment_counts[c.id] = cnt

        my_enrollments = conn.execute(
            select(enrollments.c.class_id).where(enrollments.c.student_id == student_id)
        ).fetchall()
        my_class_ids = {row.class_id for row in my_enrollments}

        student_schedule = {}
        for c in published_classes:
            if c.id in my_class_ids:
                for i in range(c.duration):
                    tiet = c.start_session + i
                    key = f"{c.day_of_week}_{c.session_type}_{tiet}"
                    if i == 0:
                        student_schedule[key] = {"status": "start", "cls": c, "rowspan": c.duration}
                    else:
                        student_schedule[key] = {"status": "blocked"}

    student_reg_open = get_setting("student_reg_open", "0") == "1"
    return render_template(
        "student/dashboard.html",
        student=student_row,
        published_classes=published_classes,
        enrollment_counts=enrollment_counts,
        my_class_ids=my_class_ids,
        student_schedule=student_schedule,
        student_reg_open=student_reg_open,
        day_name=day_name,
        session_label=session_label,
    )


@app.route("/student/enroll", methods=["POST"])
@student_required
def student_enroll():
    if get_setting("student_reg_open", "0") != "1":
        return jsonify(ok=False, error="Chưa mở đăng ký.")

    data = request.get_json(force=True)
    student_id = session["user_id"]
    student_grade = session.get("grade")

    try:
        class_id = int(data["class_id"])
    except (KeyError, ValueError, TypeError):
        return jsonify(ok=False, error="Dữ liệu không hợp lệ.")

    with engine.connect() as conn:
        cls = conn.execute(
            select(classes).where(classes.c.id == class_id)
        ).fetchone()

        if not cls:
            return jsonify(ok=False, error="Không tìm thấy lớp.")
        if not cls.is_published:
            return jsonify(ok=False, error="Lớp chưa được mở đăng ký.")
        if cls.grade != student_grade:
            return jsonify(ok=False, error="Lớp không thuộc khối của bạn.")
        if not cls.location or not cls.max_capacity:
            return jsonify(ok=False, error="Lớp chưa cập nhật đầy đủ thông tin.")

        # Check already enrolled
        existing = conn.execute(
            select(enrollments).where(
                and_(enrollments.c.student_id == student_id,
                     enrollments.c.class_id == class_id)
            )
        ).fetchone()
        if existing:
            return jsonify(ok=False, error="Bạn đã đăng ký lớp này rồi.")

        # Check capacity
        if cls.max_capacity is not None:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == class_id)
            ).scalar()
            if cnt >= cls.max_capacity:
                return jsonify(ok=False, error="Lớp đã đầy.")

    # Check time conflict
    conflict = _time_conflict(student_id, {
        "day_of_week": cls.day_of_week,
        "session_type": cls.session_type,
        "start_session": cls.start_session,
        "duration": cls.duration,
    })
    if conflict:
        with engine.connect() as conn:
            conflict_info = conn.execute(
                select(
                    classes.c.subject,
                    classes.c.day_of_week,
                    classes.c.session_type,
                    classes.c.start_session,
                    classes.c.duration,
                    teachers.c.full_name.label("teacher_name"),
                    teachers.c.subject_group,
                )
                .join(teachers, classes.c.teacher_id == teachers.c.id)
                .where(classes.c.id == conflict.id)
            ).fetchone()
        subject = (conflict_info.subject or conflict_info.subject_group) if conflict_info else f"Lớp {conflict.id}"
        teacher = conflict_info.teacher_name if conflict_info else ""
        return jsonify(
            ok=False,
            error=f"Lịch của bạn đã bị trùng với môn {subject}.",
            conflict_subject=subject,
            conflict_teacher=teacher,
            conflict_schedule=session_label(
                conflict_info.session_type,
                conflict_info.start_session,
                conflict_info.duration,
            ) + f" – {day_name(conflict_info.day_of_week)}" if conflict_info else "",
        )

    with engine.connect() as conn:
        conn.execute(
            insert(enrollments).values(
                student_id=student_id,
                class_id=class_id,
                enrolled_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
        conn.commit()

    return jsonify(ok=True)


@app.route("/student/enroll/<int:class_id>", methods=["DELETE"])
@student_required
def student_cancel_enroll(class_id):
    student_id = session["user_id"]
    with engine.connect() as conn:
        conn.execute(
            delete(enrollments).where(
                and_(enrollments.c.student_id == student_id,
                     enrollments.c.class_id == class_id)
            )
        )
        conn.commit()
    return jsonify(ok=True)


@app.route("/api/class-counts")
def api_class_counts():
    if not session.get("user_type"):
        return jsonify({}), 401

    user_type = session.get("user_type")
    with engine.connect() as conn:
        if user_type == "student":
            student_grade = session.get("grade")
            where_clause = and_(
                classes.c.is_published == 1,
                classes.c.grade == student_grade,
                classes.c.location.isnot(None),
                classes.c.location != '',
                classes.c.max_capacity.isnot(None),
            )
        elif user_type == "teacher":
            teacher_id = session.get("user_id")
            where_clause = classes.c.teacher_id == teacher_id
        else:
            return jsonify({})

        pub_classes = conn.execute(
            select(classes.c.id, classes.c.location, classes.c.max_capacity)
            .where(where_clause)
        ).fetchall()
        counts = {}
        for row in pub_classes:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == row.id)
            ).scalar()
            counts[str(row.id)] = {
                "count": cnt,
                "location": row.location or "",
                "max_capacity": row.max_capacity,
            }
    return jsonify(counts)

# ---------------------------------------------------------------------------
# Admin routes
# ---------------------------------------------------------------------------

@app.route("/admin/login", methods=["GET"])
def admin_login_page():
    return render_template("admin/login.html")


@app.route("/admin/login", methods=["POST"])
def admin_login():
    password = request.form.get("password", "")
    admin_pw = get_setting("admin_password", "admin123")
    if password == admin_pw:
        session["is_admin"] = True
        return redirect(url_for("admin_index"))
    flash("Sai mật khẩu.")
    return redirect(url_for("admin_login_page"))


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login_page"))


@app.route("/admin")
@admin_required
def admin_index():
    with engine.connect() as conn:
        teacher_count = conn.execute(select(func.count()).select_from(teachers)).scalar()
        student_count = conn.execute(select(func.count()).select_from(students)).scalar()
        class_count = conn.execute(select(func.count()).select_from(classes)).scalar()
        enrollment_count = conn.execute(select(func.count()).select_from(enrollments)).scalar()

        teacher_list = conn.execute(
            select(teachers).order_by(teachers.c.full_name)
        ).fetchall()
        student_list = conn.execute(
            select(students).order_by(students.c.grade, students.c.class_name, students.c.full_name)
        ).fetchall()

    teacher_reg_open = get_setting("teacher_reg_open", "0") == "1"
    student_reg_open = get_setting("student_reg_open", "0") == "1"
    return render_template(
        "admin/index.html",
        stats={
            "teachers": teacher_count,
            "students": student_count,
            "classes": class_count,
            "enrollments": enrollment_count,
        },
        teacher_list=teacher_list,
        student_list=student_list,
        teacher_reg_open=teacher_reg_open,
        student_reg_open=student_reg_open,
    )

# --- Admin: Teacher management ---

@app.route("/admin/teachers/upload", methods=["POST"])
@admin_required
def admin_teachers_upload():
    file = request.files.get("file")
    if not file:
        flash("Chưa chọn file.")
        return redirect(url_for("admin_index"))

    wb = openpyxl.load_workbook(file)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]

    required = ["Họ và tên", "CCCD", "Giới tính", "Tổ bộ môn"]
    for r in required:
        if r not in headers:
            flash(f"Thiếu cột: {r}")
            return redirect(url_for("admin_index"))

    idx = {h: headers.index(h) for h in required}
    count = 0
    with engine.connect() as conn:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            full_name = str(row[idx["Họ và tên"]] or "").strip()
            cccd = str(row[idx["CCCD"]] or "").strip()
            gender = str(row[idx["Giới tính"]] or "").strip()
            subject_group = str(row[idx["Tổ bộ môn"]] or "").strip()
            if not full_name or not cccd:
                continue

            existing = conn.execute(
                select(teachers).where(teachers.c.cccd == cccd)
            ).fetchone()
            if existing:
                conn.execute(
                    update(teachers).where(teachers.c.cccd == cccd).values(
                        full_name=full_name,
                        gender=gender,
                        subject_group=subject_group,
                    )
                )
            else:
                conn.execute(
                    insert(teachers).values(
                        full_name=full_name,
                        cccd=cccd,
                        gender=gender,
                        subject_group=subject_group,
                        email=None,
                        password_hash=None,
                        is_first_login=1,
                    )
                )
            count += 1
        conn.commit()

    flash(f"Đã nhập {count} giáo viên.")
    return redirect(url_for("admin_index"))


@app.route("/admin/teachers/template")
@admin_required
def admin_teachers_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Giáo viên"
    ws.append(["Họ và tên", "CCCD", "Giới tính", "Tổ bộ môn"])
    ws.append(["Nguyễn Văn A", "012345678901", "Nam", "Toán"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="mau_giao_vien.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/teachers/add", methods=["POST"])
@admin_required
def admin_teachers_add():
    full_name = request.form.get("full_name", "").strip()
    cccd = request.form.get("cccd", "").strip()
    gender = request.form.get("gender", "").strip()
    subject_group = request.form.get("subject_group", "").strip()

    if not full_name or not cccd:
        flash("Vui lòng nhập đầy đủ thông tin.")
        return redirect(url_for("admin_index"))

    with engine.connect() as conn:
        existing = conn.execute(
            select(teachers).where(teachers.c.cccd == cccd)
        ).fetchone()
        if existing:
            conn.execute(
                update(teachers).where(teachers.c.cccd == cccd).values(
                    full_name=full_name,
                    gender=gender,
                    subject_group=subject_group,
                )
            )
            flash("Đã cập nhật giáo viên.")
        else:
            conn.execute(
                insert(teachers).values(
                    full_name=full_name,
                    cccd=cccd,
                    gender=gender,
                    subject_group=subject_group,
                    email=None,
                    password_hash=None,
                    is_first_login=1,
                )
            )
            flash("Đã thêm giáo viên.")
        conn.commit()

    return redirect(url_for("admin_index"))


@app.route("/admin/teachers/<int:teacher_id>/reset", methods=["POST"])
@admin_required
def admin_teachers_reset(teacher_id):
    with engine.connect() as conn:
        # Get all classes for this teacher
        teacher_classes = conn.execute(
            select(classes.c.id).where(classes.c.teacher_id == teacher_id)
        ).fetchall()
        class_ids = [c.id for c in teacher_classes]

        # Delete enrollments for those classes
        if class_ids:
            conn.execute(
                delete(enrollments).where(enrollments.c.class_id.in_(class_ids))
            )
            conn.execute(
                delete(classes).where(classes.c.teacher_id == teacher_id)
            )

        # Reset teacher account
        conn.execute(
            update(teachers).where(teachers.c.id == teacher_id).values(
                email=None,
                password_hash=None,
                is_first_login=1,
            )
        )
        conn.commit()

    flash("Đã reset tài khoản giáo viên.")
    return redirect(url_for("admin_index"))


@app.route("/admin/teachers/<int:teacher_id>/reset-password", methods=["POST"])
@admin_required
def admin_teachers_reset_password(teacher_id):
    with engine.connect() as conn:
        teacher = conn.execute(
            select(teachers.c.id, teachers.c.full_name, teachers.c.is_first_login)
            .where(teachers.c.id == teacher_id)
        ).fetchone()
    if not teacher:
        return jsonify(ok=False, error="Không tìm thấy giáo viên.")
    if teacher.is_first_login:
        return jsonify(ok=False, error="Tài khoản chưa kích hoạt, không cần khôi phục mật khẩu.")

    temp_pw = generate_temp_password()
    with engine.begin() as conn:
        conn.execute(
            update(teachers).where(teachers.c.id == teacher_id).values(
                password_hash=generate_password_hash(temp_pw),
                must_change_password=1,
            )
        )
    return jsonify(ok=True, new_password=temp_pw, name=teacher.full_name)


@app.route("/admin/teachers/<int:teacher_id>", methods=["DELETE"])
@admin_required
def admin_teachers_delete(teacher_id):
    with engine.connect() as conn:
        teacher_classes = conn.execute(
            select(classes.c.id).where(classes.c.teacher_id == teacher_id)
        ).fetchall()
        class_ids = [c.id for c in teacher_classes]

        if class_ids:
            conn.execute(
                delete(enrollments).where(enrollments.c.class_id.in_(class_ids))
            )
            conn.execute(
                delete(classes).where(classes.c.teacher_id == teacher_id)
            )

        conn.execute(delete(teachers).where(teachers.c.id == teacher_id))
        conn.commit()

    return jsonify(ok=True)

@app.route("/admin/teachers/clear", methods=["POST"])
@admin_required
def admin_teachers_clear():
    with engine.begin() as conn:
        all_class_ids = [r.id for r in conn.execute(select(classes.c.id)).fetchall()]
        if all_class_ids:
            conn.execute(delete(enrollments).where(enrollments.c.class_id.in_(all_class_ids)))
        conn.execute(delete(classes))
        conn.execute(delete(teachers))
    flash("Đã xóa toàn bộ danh sách giáo viên.", "success")
    return redirect(url_for("admin_index"))


# --- Admin: Student management ---

@app.route("/admin/students/upload", methods=["POST"])
@admin_required
def admin_students_upload():
    file = request.files.get("file")
    if not file:
        flash("Chưa chọn file.")
        return redirect(url_for("admin_index"))

    wb = openpyxl.load_workbook(file)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]

    required = ["Họ và tên", "CCCD", "Lớp", "Khối"]
    for r in required:
        if r not in headers:
            flash(f"Thiếu cột: {r}")
            return redirect(url_for("admin_index"))

    idx = {h: headers.index(h) for h in required}
    count = 0
    with engine.connect() as conn:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            full_name = str(row[idx["Họ và tên"]] or "").strip()
            cccd = str(row[idx["CCCD"]] or "").strip()
            class_name = str(row[idx["Lớp"]] or "").strip()
            try:
                grade = int(row[idx["Khối"]] or 0)
            except (ValueError, TypeError):
                grade = 0
            if not full_name or not cccd:
                continue

            existing = conn.execute(
                select(students).where(students.c.cccd == cccd)
            ).fetchone()
            if existing:
                conn.execute(
                    update(students).where(students.c.cccd == cccd).values(
                        full_name=full_name,
                        class_name=class_name,
                        grade=grade,
                    )
                )
            else:
                conn.execute(
                    insert(students).values(
                        full_name=full_name,
                        cccd=cccd,
                        class_name=class_name,
                        grade=grade,
                    )
                )
            count += 1
        conn.commit()

    flash(f"Đã nhập {count} học sinh.")
    return redirect(url_for("admin_index"))


@app.route("/admin/students/template")
@admin_required
def admin_students_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Học sinh"
    ws.append(["Họ và tên", "CCCD", "Lớp", "Khối"])
    ws.append(["Trần Thị B", "098765432109", "10A1", 10])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="mau_hoc_sinh.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/students/add", methods=["POST"])
@admin_required
def admin_students_add():
    full_name = request.form.get("full_name", "").strip()
    cccd = request.form.get("cccd", "").strip()
    class_name = request.form.get("class_name", "").strip()
    try:
        grade = int(request.form.get("grade", 0))
    except ValueError:
        grade = 0

    if not full_name or not cccd:
        flash("Vui lòng nhập đầy đủ thông tin.")
        return redirect(url_for("admin_index"))

    with engine.connect() as conn:
        existing = conn.execute(
            select(students).where(students.c.cccd == cccd)
        ).fetchone()
        if existing:
            conn.execute(
                update(students).where(students.c.cccd == cccd).values(
                    full_name=full_name,
                    class_name=class_name,
                    grade=grade,
                )
            )
            flash("Đã cập nhật học sinh.")
        else:
            conn.execute(
                insert(students).values(
                    full_name=full_name,
                    cccd=cccd,
                    class_name=class_name,
                    grade=grade,
                )
            )
            flash("Đã thêm học sinh.")
        conn.commit()

    return redirect(url_for("admin_index"))


@app.route("/admin/students/<int:student_id>/reset", methods=["POST"])
@admin_required
def admin_students_reset(student_id):
    with engine.begin() as conn:
        conn.execute(delete(enrollments).where(enrollments.c.student_id == student_id))
        conn.execute(
            update(students).where(students.c.id == student_id).values(
                password_hash=None, is_first_login=1
            )
        )
    flash("Đã reset tài khoản và xóa đăng ký của học sinh.", "success")
    return redirect(url_for("admin_index"))


@app.route("/admin/students/<int:student_id>/reset-password", methods=["POST"])
@admin_required
def admin_students_reset_password(student_id):
    with engine.connect() as conn:
        student = conn.execute(
            select(students.c.id, students.c.full_name, students.c.is_first_login)
            .where(students.c.id == student_id)
        ).fetchone()
    if not student:
        return jsonify(ok=False, error="Không tìm thấy học sinh.")
    if student.is_first_login:
        return jsonify(ok=False, error="Tài khoản chưa kích hoạt, không cần khôi phục mật khẩu.")

    temp_pw = generate_temp_password()
    with engine.begin() as conn:
        conn.execute(
            update(students).where(students.c.id == student_id).values(
                password_hash=generate_password_hash(temp_pw),
                must_change_password=1,
            )
        )
    return jsonify(ok=True, new_password=temp_pw, name=student.full_name)


@app.route("/admin/students/<int:student_id>", methods=["DELETE"])
@admin_required
def admin_students_delete(student_id):
    with engine.connect() as conn:
        conn.execute(
            delete(enrollments).where(enrollments.c.student_id == student_id)
        )
        conn.execute(delete(students).where(students.c.id == student_id))
        conn.commit()
    return jsonify(ok=True)


@app.route("/admin/students/clear", methods=["POST"])
@admin_required
def admin_students_clear():
    with engine.begin() as conn:
        conn.execute(delete(enrollments))
        conn.execute(delete(students))
    flash("Đã xóa toàn bộ danh sách học sinh.", "success")
    return redirect(url_for("admin_index"))


# --- Admin: Class registration management ---

@app.route("/admin/class-reg")
@admin_required
def admin_class_reg():
    with engine.connect() as conn:
        all_teachers = conn.execute(
            select(teachers).order_by(teachers.c.subject_group, teachers.c.full_name)
        ).fetchall()

        all_classes = conn.execute(
            select(classes, teachers.c.full_name.label("teacher_name"),
                   teachers.c.subject_group, teachers.c.email.label("teacher_email"))
            .join(teachers, classes.c.teacher_id == teachers.c.id)
            .order_by(teachers.c.subject_group, teachers.c.full_name, classes.c.start_session)
        ).fetchall()

        enrollment_counts = {}
        for c in all_classes:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == c.id)
            ).scalar()
            enrollment_counts[c.id] = cnt

    teacher_reg_open = get_setting("teacher_reg_open", "0") == "1"
    return render_template(
        "admin/class_reg.html",
        all_teachers=all_teachers,
        all_classes=all_classes,
        enrollment_counts=enrollment_counts,
        teacher_reg_open=teacher_reg_open,
        day_name=day_name,
        session_label=session_label,
    )


@app.route("/admin/class-reg/export")
@admin_required
def admin_class_reg_export():
    with engine.connect() as conn:
        all_classes = conn.execute(
            select(classes, teachers.c.full_name.label("teacher_name"),
                   teachers.c.subject_group, teachers.c.email.label("teacher_email"))
            .join(teachers, classes.c.teacher_id == teachers.c.id)
            .order_by(classes.c.grade, classes.c.day_of_week,
                      classes.c.session_type, classes.c.start_session)
        ).fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Phân phòng học"
    headers = ["id", "Họ và tên GV", "Tổ bộ môn", "Khối", "Thứ",
               "Buổi", "Tiết BĐ", "Số tiết", "Môn học", "Địa điểm", "Sĩ số"]
    ws.append(headers)
    # Header style
    from openpyxl.styles import PatternFill, Font
    hdr_fill = PatternFill("solid", fgColor="0369A1")
    hdr_font = Font(color="FFFFFF", bold=True)
    for col_idx, _ in enumerate(headers, 1):
        cell = ws.cell(1, col_idx)
        cell.fill = hdr_fill
        cell.font = hdr_font
    for c in all_classes:
        buoi = "Sáng" if c.session_type == "morning" else "Chiều"
        ws.append([
            c.id,
            c.teacher_name,
            c.subject_group,
            c.grade,
            day_name(c.day_of_week),
            buoi,
            c.start_session,
            c.duration,
            c.subject or "",
            c.location or "",
            c.max_capacity or "",
        ])
    # Column widths
    for col, width in zip([1,2,3,4,5,6,7,8,9,10,11],
                          [6,22,14,6,8,7,8,8,18,14,8]):
        ws.column_dimensions[get_column_letter(col)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="phan_phong_hoc.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/classes/import-rooms", methods=["POST"])
@admin_required
def admin_classes_import_rooms():
    file = request.files.get("file")
    if not file:
        flash("Chưa chọn file.")
        return redirect(url_for("admin_class_reg"))
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    if "id" not in headers or "Địa điểm" not in headers or "Sĩ số" not in headers:
        flash("File thiếu cột: cần có 'id', 'Địa điểm', 'Sĩ số'.")
        return redirect(url_for("admin_class_reg"))
    count = 0
    with engine.connect() as conn:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            try:
                class_id = int(row[headers.index("id")] or 0)
            except (ValueError, TypeError):
                continue
            location = str(row[headers.index("Địa điểm")] or "").strip() or None
            try:
                max_capacity = int(row[headers.index("Sĩ số")] or 0) or None
            except (ValueError, TypeError):
                max_capacity = None
            conn.execute(
                update(classes).where(classes.c.id == class_id).values(
                    location=location, max_capacity=max_capacity
                )
            )
            count += 1
        conn.commit()
    flash(f"Đã cập nhật phòng học cho {count} lớp.")
    return redirect(url_for("admin_class_reg"))


@app.route("/admin/classes/<int:class_id>", methods=["PATCH"])
@admin_required
def admin_class_update(class_id):
    data = request.get_json(force=True, silent=True) or {}
    location = (data.get("location") or "").strip() or None
    try:
        max_capacity = int(data["max_capacity"]) if data.get("max_capacity") else None
    except (ValueError, TypeError):
        max_capacity = None
    with engine.begin() as conn:
        conn.execute(
            update(classes).where(classes.c.id == class_id).values(
                location=location, max_capacity=max_capacity
            )
        )
    return jsonify(ok=True, location=location, max_capacity=max_capacity)


@app.route("/admin/class-reg/toggle", methods=["POST"])
@admin_required
def admin_class_reg_toggle():
    current = get_setting("teacher_reg_open", "0")
    new_val = "0" if current == "1" else "1"
    set_setting("teacher_reg_open", new_val)
    return jsonify(ok=True, open=new_val == "1")


@app.route("/admin/classes/<int:class_id>/publish", methods=["POST"])
@admin_required
def admin_class_publish(class_id):
    data = request.get_json(force=True, silent=True) or {}
    publish = 1 if data.get("publish") else 0
    with engine.begin() as conn:
        conn.execute(
            update(classes).where(classes.c.id == class_id).values(is_published=publish)
        )
    return jsonify(ok=True, is_published=bool(publish))


@app.route("/admin/classes/<int:class_id>", methods=["DELETE"])
@admin_required
def admin_class_delete(class_id):
    with engine.begin() as conn:
        conn.execute(delete(enrollments).where(enrollments.c.class_id == class_id))
        conn.execute(delete(classes).where(classes.c.id == class_id))
    return jsonify(ok=True)


# --- Admin: Enrollment management ---

@app.route("/admin/enrollment")
@admin_required
def admin_enrollment():
    with engine.connect() as conn:
        published_classes = conn.execute(
            select(classes, teachers.c.full_name.label("teacher_name"),
                   teachers.c.subject_group)
            .join(teachers, classes.c.teacher_id == teachers.c.id)
            .where(classes.c.is_published == 1)
            .order_by(classes.c.grade, classes.c.day_of_week,
                      classes.c.session_type, classes.c.start_session)
        ).fetchall()

        enrollment_counts = {}
        for c in published_classes:
            cnt = conn.execute(
                select(func.count()).where(enrollments.c.class_id == c.id)
            ).scalar()
            enrollment_counts[c.id] = cnt

    # Collect extra_data column names across all published classes
    extra_columns = set()
    for c in published_classes:
        if c.extra_data:
            try:
                d = json.loads(c.extra_data)
                extra_columns.update(d.keys())
            except Exception:
                pass

    student_reg_open = get_setting("student_reg_open", "0") == "1"
    return render_template(
        "admin/enrollment.html",
        published_classes=published_classes,
        enrollment_counts=enrollment_counts,
        extra_columns=sorted(extra_columns),
        student_reg_open=student_reg_open,
        day_name=day_name,
        session_label=session_label,
    )


@app.route("/admin/enrollment/upload", methods=["POST"])
@admin_required
def admin_enrollment_upload():
    file = request.files.get("file")
    if not file:
        flash("Chưa chọn file.")
        return redirect(url_for("admin_enrollment"))

    wb = openpyxl.load_workbook(file)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]

    required = ["id", "Địa điểm", "Sĩ số"]
    for r in required:
        if r not in headers:
            flash(f"Thiếu cột: {r}")
            return redirect(url_for("admin_enrollment"))

    extra_cols = [h for h in headers if h not in required and h]
    count = 0
    with engine.connect() as conn:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            try:
                class_id = int(row[headers.index("id")] or 0)
            except (ValueError, TypeError):
                continue
            location = str(row[headers.index("Địa điểm")] or "").strip()
            try:
                max_capacity = int(row[headers.index("Sĩ số")] or 0)
            except (ValueError, TypeError):
                max_capacity = None

            extra = {}
            for col in extra_cols:
                val = row[headers.index(col)]
                if val is not None:
                    extra[col] = str(val)

            extra_data_json = json.dumps(extra, ensure_ascii=False) if extra else None

            conn.execute(
                update(classes).where(classes.c.id == class_id).values(
                    location=location,
                    max_capacity=max_capacity,
                    extra_data=extra_data_json,
                    is_published=1,
                )
            )
            count += 1
        conn.commit()

    flash(f"Đã cập nhật {count} lớp và mở đăng ký cho học sinh.")
    return redirect(url_for("admin_enrollment"))


@app.route("/admin/enrollment/add", methods=["POST"])
@admin_required
def admin_enrollment_add():
    data = request.get_json(force=True)
    try:
        teacher_id = int(data["teacher_id"])
        grade = int(data["grade"])
        day_of_week = int(data["day_of_week"])
        session_type = data["session_type"]
        start_session = int(data["start_session"])
        duration = int(data["duration"])
        subject = (data.get("subject") or "").strip() or None
        location = (data.get("location") or "").strip() or None
        max_capacity = int(data["max_capacity"]) if data.get("max_capacity") else None
        extra_data_json = data.get("extra_data_json") or None
    except (KeyError, ValueError, TypeError) as e:
        return jsonify(ok=False, error=f"Dữ liệu không hợp lệ: {e}")

    with engine.connect() as conn:
        result = conn.execute(
            insert(classes).values(
                teacher_id=teacher_id,
                grade=grade,
                duration=duration,
                day_of_week=day_of_week,
                session_type=session_type,
                start_session=start_session,
                subject=subject,
                location=location,
                max_capacity=max_capacity,
                extra_data=extra_data_json,
                is_published=1,
                created_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
        conn.commit()
        class_id = result.inserted_primary_key[0]

    return jsonify(ok=True, class_id=class_id)


@app.route("/admin/enrollment/export")
@admin_required
def admin_enrollment_export():
    with engine.connect() as conn:
        published_classes = conn.execute(
            select(classes, teachers.c.full_name.label("teacher_name"),
                   teachers.c.subject_group, teachers.c.email.label("teacher_email"))
            .join(teachers, classes.c.teacher_id == teachers.c.id)
            .where(classes.c.is_published == 1)
            .order_by(classes.c.grade, classes.c.day_of_week,
                      classes.c.session_type, classes.c.start_session)
        ).fetchall()

    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    summary_rows = []

    with engine.connect() as conn:
        for cls in published_classes:
            sheet_name = f"Lớp {cls.id}: {cls.subject or 'N/A'} K{cls.grade}"
            # Excel sheet names max 31 chars
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["STT", "Họ tên", "Lớp", "Khối"])

            enrolled_students = conn.execute(
                select(students, enrollments.c.enrolled_at)
                .join(enrollments, students.c.id == enrollments.c.student_id)
                .where(enrollments.c.class_id == cls.id)
                .order_by(students.c.class_name, students.c.full_name)
            ).fetchall()

            for i, s in enumerate(enrolled_students, 1):
                ws.append([i, s.full_name, s.class_name, s.grade])

            enrolled_count = len(enrolled_students)
            buoi = "Sáng" if cls.session_type == "morning" else "Chiều"
            summary_rows.append([
                f"Lớp {cls.id}",
                cls.teacher_name,
                cls.subject or "",
                cls.grade,
                day_name(cls.day_of_week),
                buoi,
                cls.start_session,
                cls.location or "",
                enrolled_count,
                cls.max_capacity or "",
            ])

    # Summary sheet
    ws_sum = wb.create_sheet(title="Tổng hợp", index=0)
    ws_sum.append(["Lớp", "GV", "Môn", "Khối", "Thứ", "Buổi", "Tiết",
                   "Địa điểm", "Sĩ số đã đăng ký", "Sĩ số tối đa"])
    for row in summary_rows:
        ws_sum.append(row)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="danh_sach_dang_ky.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/enrollment/toggle", methods=["POST"])
@admin_required
def admin_enrollment_toggle():
    current = get_setting("student_reg_open", "0")
    new_val = "0" if current == "1" else "1"
    set_setting("student_reg_open", new_val)
    return jsonify(ok=True, open=new_val == "1")


@app.route("/admin/enrollment/<int:class_id>/students")
@admin_required
def admin_enrollment_students(class_id):
    with engine.connect() as conn:
        enrolled = conn.execute(
            select(students, enrollments.c.enrolled_at)
            .join(enrollments, students.c.id == enrollments.c.student_id)
            .where(enrollments.c.class_id == class_id)
            .order_by(students.c.class_name, students.c.full_name)
        ).fetchall()

    return jsonify(ok=True, students=[
        {
            "id": s.id,
            "full_name": s.full_name,
            "class_name": s.class_name,
            "grade": s.grade,
            "enrolled_at": s.enrolled_at,
        }
        for s in enrolled
    ])

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True, port=5051)
