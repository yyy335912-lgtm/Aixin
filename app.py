"""浙江大学学生爱心社 积分系统 - Flask 主程序"""
import os
import io
import csv
import json
import re
import random
import string
from datetime import datetime, timedelta
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    session, jsonify, send_file, abort,
)
from sqlalchemy import func, desc

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from models import (
    db, Member, Project, Activity, ActivityParticipant,
    PositionRecord, IrregularActivity, IrregularParticipant,
    QueryLog, SignoffLog, ModificationLog, Blacklist, Setting,
    LongTermGroup, Admin,
)


def create_app():
    app = Flask(__name__)
    db_path = os.path.join(os.path.dirname(__file__), "data", "aixin.db")
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "connect_args": {"check_same_thread": False},
        "pool_pre_ping": True,
        "pool_recycle": 300,
    }
    app.config["SECRET_KEY"] = "zju-aixin-secret-2026"
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

    # 注册 Jinja 全局：_has(perm) 用于侧边栏按权限过滤
    @app.context_processor
    def _inject_perm_helpers():
        def _has(perm):
            me_id = session.get("admin_id")
            if not me_id:
                return False
            me = Admin.query.get(me_id)
            if me is None:
                return False
            if me.is_super_admin:
                return True
            return me.has_permission(perm)
        def _perm_label(me):
            perms = me.permissions_list() if me else []
            return f"{len(perms)}项" if perms else "无"
        return dict(_has=_has, _perm_label=_perm_label, me=Admin.query.get(session.get("admin_id")))
    db.init_app(app)

    # SQLite 并发优化：WAL 模式 + 忙等待超时
    with app.app_context():
        try:
            from sqlalchemy import event
            @event.listens_for(db.engine, "connect")
            def _set_sqlite_pragma(dbapi_connection, connection_record):
                cursor = dbapi_connection.cursor()
                cursor.execute("PRAGMA journal_mode=WAL")
                cursor.execute("PRAGMA busy_timeout=5000")
                cursor.execute("PRAGMA synchronous=NORMAL")
                cursor.execute("PRAGMA cache_size=-8000")
                cursor.execute("PRAGMA temp_store=MEMORY")
                cursor.close()
        except Exception:
            pass

    register_routes(app)
    return app


# ============== 辅助函数 ==============

# 所有可分配的管理权限（与 models.AllAdminPermissions 保持同步）
# 白名单管理（whitelist）仅超管使用，不在此列
ADMIN_PERMISSIONS = [
    ("dashboard",       "🧡 后台首页",      "查看统计面板与快捷入口"),
    ("members",         "👥 成员管理",      "增删改查社员、批量导入导出"),
    ("projects",        "🌟 项目管理",      "管理心暖系列等品牌项目"),
    ("activities",      "📖 活动管理",      "登记/编辑活动，导入参与名单"),
    ("positions",       "🎖️ 任职积分",      "录入/导入部长副部长等任职积分"),
    ("irregulars",      "✦ 非常规活动",     "唐奖、年检、招新等零散事件"),
    ("long_term_groups","🧡 长期项目组",    "管理 7 大长期项目组"),
    ("blacklist",       "🚫 黑名单",       "加入/移出志愿者黑名单"),
    ("stats",           "📈 数据统计",      "查看积分榜、部门/项目汇总"),
    ("logs",            "📋 审计日志",      "查询/签收/修改记录"),
    ("export",          "📤 数据导出",      "导出 Excel 全量数据"),
    ("settings",        "🔧 系统设置",      "学期/部门等基本配置"),
    ("profile",         "⚙️ 我的资料",      "修改个人信息、密码"),
]


def admin_required(f):
    """仅检查登录态（任何 admin 都可访问）"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("admin_id"):
            return redirect(url_for("admin_login", next=request.path))
        return f(*args, **kwargs)
    return wrapper


def permission_required(perm):
    """检查登录 + 是否有某项权限
    - 超管自动通过
    - 普通管理员需 permissions 字段中包含 perm
    """
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not session.get("admin_id"):
                return redirect(url_for("admin_login", next=request.path))
            me = Admin.query.get(session.get("admin_id"))
            if not me:
                session.clear()
                return redirect(url_for("admin_login"))
            if not me.has_permission(perm):
                label = next((n for k, n, _ in ADMIN_PERMISSIONS if k == perm), perm)
                flash(f"您没有「{label}」的访问权限", "warning")
                return redirect(url_for("admin_login"))
            return f(*args, **kwargs)
        return wrapper
    return decorator


def super_admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("admin_id"):
            return redirect(url_for("admin_login", next=request.path))
        if not session.get("is_super_admin"):
            flash("此操作仅限超级管理员", "danger")
            return redirect(url_for("admin_dashboard"))
        return f(*args, **kwargs)
    return wrapper


def current_semester():
    s = Setting.query.filter_by(key="current_semester").first()
    return s.value if s else ""


def normalize_grade(raw):
    """统一年级格式：保留纯数字部分，追加 '级'"""
    s = (raw or "").strip().replace("级", "")
    return s + "级" if s else raw


def all_departments():
    s = Setting.query.filter_by(key="departments").first()
    if s and s.value:
        return [d.strip() for d in s.value.split(",") if d.strip()]
    return ["人资部", "综管部", "文体部", "活动部",
            "宣传部", "创设部", "外联部", "财务部"]


def get_setting(key, default=""):
    s = Setting.query.filter_by(key=key).first()
    return s.value if s else default


def log_modification(modifier, member_id, details, query_code=""):
    db.session.add(ModificationLog(
        modifier=modifier, member_id=member_id,
        details=details, query_code=query_code,
    ))


def detect_current_semester():
    """根据浙大学历制自动检测当前学年和学期
    秋冬学期：9月-次年1月
    春夏学期：2月-8月
    返回 (学年字符串如 '2025-2026', 季节字符串如 '秋冬')
    """
    from datetime import datetime
    m = datetime.now().month
    y = datetime.now().year
    if m >= 9:
        return f"{y}-{y+1}", "秋冬"
    elif m >= 2:
        return f"{y-1}-{y}", "春夏"
    else:
        return f"{y-1}-{y}", "秋冬"


# ============== 路由注册 ==============

def register_routes(app):

    @app.context_processor
    def inject_globals():
        def _pair_chunks(s, n=2):
            """把字符串每 n 字一组，列表返回（用于 2 列字符布局）
            >>> _pair_chunks('心暖陇原', 2) == ['心暖', '陇原']
            >>> _pair_chunks('爱教三部曲', 2) == ['爱教', '三部', '曲']
            """
            s = s or ""
            return [s[i:i+n] for i in range(0, len(s), n)]
        return {
            "now": datetime.now(),
            "current_semester": current_semester(),
            "departments": all_departments(),
            "_pair_chunks": _pair_chunks,
        }

    # ---------- 公共首页 ----------
    @app.route("/")
    def index():
        a_hours = db.session.query(func.sum(ActivityParticipant.a_hours)).scalar() or 0
        a_ir = db.session.query(func.sum(IrregularParticipant.a_hours)).scalar() or 0
        b_hours = db.session.query(func.sum(ActivityParticipant.b_hours)).scalar() or 0
        b_ir = db.session.query(func.sum(IrregularParticipant.b_hours)).scalar() or 0
        stats = {
            "members": Member.query.count(),
            "projects": Project.query.filter_by(active=True).count(),
            "activities": Activity.query.count(),
            "a_hours": round(float(a_hours) + float(a_ir), 1),
            "b_hours": round(float(b_hours) + float(b_ir), 1),
        }
        # 积分龙虎榜 Top 10
        top = []
        for m in Member.query.all():
            top.append({"name": m.name, "points": m.total_points(), "grade": m.grade, "dept": m.department})
        top.sort(key=lambda x: -x["points"])
        top = top[:10]
        # 近期活动（最近 8 场）
        recent = Activity.query.order_by(Activity.id.desc()).limit(8).all()
        return render_template("index.html", stats=stats, top=top, recent=recent)

    # ---------- 社员自助查询（学号 + 姓名）----------
    @app.route("/query", methods=["GET", "POST"])
    def query():
        if request.method == "POST":
            sid = (request.form.get("student_id") or "").strip()
            name = (request.form.get("queryer_name") or "").strip()
            if not sid:
                flash("请输入学号", "warning")
                return redirect(url_for("query"))
            member = Member.query.filter_by(student_id=sid).first()
            # 记录查询
            db.session.add(QueryLog(
                member_id=member.id if member else None,
                queryer_name=name, query_code=sid,
            ))
            db.session.commit()
            if not member:
                flash("未找到该学号对应的社员，请检查输入是否正确", "danger")
                return redirect(url_for("query"))
            return redirect(url_for("query_result", sid=sid, name=name))
        return render_template("query.html")

    @app.route("/query/result")
    def query_result():
        sid = request.args.get("sid", "")
        name = request.args.get("name", "")
        member = Member.query.filter_by(student_id=sid).first_or_404()
        # 所有活动平铺列表（按日期降序，含累计）
        all_activities = []
        total_a = total_b = 0.0
        # 常规活动
        for ap in member.activity_participations:
            a = ap.activity
            if not a: continue
            total_a += ap.a_hours or 0
            total_b += ap.b_hours or 0
            all_activities.append({
                "type": "regular", "id": f"ap:{ap.id}",
                "date": a.date or "", "project": a.project.name if a.project else "",
                "content": a.content or "", "location": a.location or "",
                "a_hours": ap.a_hours or 0, "b_hours": ap.b_hours or 0,
                "points": ap.points(),
                "semester": a.semester or "",
                "is_signed_off": ap.is_signed_off, "signed_off_at": ap.signed_off_at,
            })
        # 非常规活动
        for ip in member.irregular_participations:
            ir = ip.irregular
            if not ir: continue
            total_a += ip.a_hours or 0
            total_b += ip.b_hours or 0
            all_activities.append({
                "type": "irregular", "id": f"ip:{ip.id}",
                "date": ir.date or "", "project": "🎲 非常规",
                "content": ir.content or "", "location": ir.location or "",
                "a_hours": ip.a_hours or 0, "b_hours": ip.b_hours or 0,
                "points": ip.points(),
                "semester": ir.semester or "",
                "is_signed_off": ip.is_signed_off, "signed_off_at": ip.signed_off_at,
            })
        all_activities.sort(key=lambda x: x["date"], reverse=True)
        # 项目分类汇总（保留给页头概要）
        project_data = []
        for p in Project.query.filter_by(active=True).order_by(Project.sort_order).all():
            a_h, b_h = member.project_points(p.id)
            project_data.append({
                "project": p,
                "a_hours": a_h,
                "b_hours": b_h,
                "points": round(a_h + b_h / 2.0, 2),
            })
        # 排名：全体 / 部门 / 长期项目组
        all_members = Member.query.all()
        sorted_all = sorted(all_members, key=lambda x: x.total_points(), reverse=True)
        rank_all = next((i + 1 for i, m in enumerate(sorted_all) if m.id == member.id), 0)
        # 部门内排名
        if member.department:
            dept_members = [m for m in all_members if m.department == member.department]
            sorted_dept = sorted(dept_members, key=lambda x: x.total_points(), reverse=True)
            rank_dept = next((i + 1 for i, m in enumerate(sorted_dept) if m.id == member.id), 0)
            dept_total = len(dept_members)
        else:
            rank_dept = 0
            dept_total = 0
        # 长期项目组内排名
        if member.long_term_group_id:
            lt_members = [m for m in all_members if m.long_term_group_id == member.long_term_group_id]
            sorted_lt = sorted(lt_members, key=lambda x: x.total_points(), reverse=True)
            rank_lt = next((i + 1 for i, m in enumerate(sorted_lt) if m.id == member.id), 0)
            lt_total = len(lt_members)
            lt_name = member.long_term_group.name
        else:
            rank_lt = 0
            lt_total = 0
            lt_name = None
        return render_template(
            "query_result.html", member=member, project_data=project_data,
            all_activities=all_activities,
            total_a=round(total_a, 1), total_b=round(total_b, 1),
            name=name,
            rank_all=rank_all, total_all=len(all_members),
            rank_dept=rank_dept, total_dept=dept_total,
            rank_lt=rank_lt, total_lt=lt_total, lt_name=lt_name,
            pending_count=member.unsigned_activity_count(),
            latest_signoff_at=member.latest_signoff_str(),
        )

    @app.route("/query/signoff", methods=["POST"])
    def query_signoff():
        """签收 - 支持全部 / 单条 / 选中多条
        表单字段:
          student_id: 必填，成员学号
          ids: 可选，勾选的参与记录 id（ap_id 或 ip_id 前缀）
          mode: 'all'（默认）/ 'selected'
        """
        sid = (request.form.get("student_id") or "").strip()
        member = Member.query.filter_by(student_id=sid).first_or_404()
        mode = request.form.get("mode", "all")
        raw_ids = request.form.get("ids", "")
        ids = [x.strip() for x in raw_ids.split(",") if x.strip()]

        now = datetime.utcnow()
        signed_count = 0

        if mode == "all":
            # 签收所有未签收的活动
            for ap in member.activity_participations:
                if ap.signed_off_at is None:
                    ap.signed_off_at = now
                    signed_count += 1
            for ip in member.irregular_participations:
                if ip.signed_off_at is None:
                    ip.signed_off_at = now
                    signed_count += 1
        else:
            # 签收选中的（ids 形如 "ap:123" / "ip:456"）
            ap_ids, ip_ids = set(), set()
            for raw in ids:
                if raw.startswith("ap:"):
                    ap_ids.add(int(raw[3:]))
                elif raw.startswith("ip:"):
                    ip_ids.add(int(raw[3:]))
            for ap in member.activity_participations:
                if ap.id in ap_ids and ap.signed_off_at is None:
                    ap.signed_off_at = now
                    signed_count += 1
            for ip in member.irregular_participations:
                if ip.id in ip_ids and ip.signed_off_at is None:
                    ip.signed_off_at = now
                    signed_count += 1

        # 写审计日志
        db.session.add(SignoffLog(
            method=f"{'签收全部' if mode == 'all' else '签收选中'}（{signed_count} 项）",
            member_id=member.id,
            related_query_code=sid,
        ))
        db.session.commit()
        if signed_count > 0:
            flash(f"【{member.name}】已签收 {signed_count} 项活动 ✅", "success")
        else:
            flash("没有需要签收的项", "info")
        return redirect(url_for("query_result", sid=sid))

    @app.route("/member/update", methods=["POST"])
    def member_update():
        """成员自助修改个人信息"""
        sid = (request.form.get("student_id") or "").strip()
        qname = (request.form.get("queryer_name") or "").strip()
        member = Member.query.filter_by(student_id=sid).first()
        if not member:
            flash("学号不存在，无法修改", "danger")
            return redirect(url_for("query_result", sid=sid))
        phone = (request.form.get("phone") or "").strip()
        grade = normalize_grade(request.form.get("grade"))
        changes = []
        if phone != (member.phone or ""):
            changes.append(f"手机号：{member.phone or '空'} → {phone}")
            member.phone = phone
        if grade != (member.grade or ""):
            changes.append(f"年级：{member.grade or '空'} → {grade}")
            member.grade = grade
        if changes:
            log_modification(name, member.id, "成员自助修改：" + "；".join(changes))
            db.session.commit()
            flash(f"✅ 信息已更新：{'；'.join(changes)}", "success")
        else:
            flash("未检测到任何修改", "info")
        return redirect(url_for("query_result", sid=sid))

    # ---------- 管理员登录 / 注册（审批流） ----------
    @app.route("/admin/login", methods=["GET", "POST"])
    def admin_login():
        """登录/完成注册
        状态机：
          - 不在 admins 表 → 拒绝登录，提示"请先申请"
          - status=pending（已申请未审批）→ 拒绝登录，提示"审批中"
          - status=approved 且无密码 → 「完成注册」：设置密码，激活
          - status=active + 密码正确 → 登录
          - status=active + 密码错 → 拒绝
        """
        if request.method == "POST":
            name = (request.form.get("name") or "").strip()
            sid = (request.form.get("student_id") or "").strip()
            pw = request.form.get("password", "")
            remember = request.form.get("remember") == "1"

            # 1. 必须是社团成员
            member = Member.query.filter_by(name=name, student_id=sid).first()
            if not member:
                flash(
                    f"❌ {name or '（空）'}（学号 {sid or '（空）'}）不在社团成员名单中，"
                    f"请联系社长/超管先将您加入「成员管理」。",
                    "danger",
                )
                return render_template("admin/login.html", name=name, student_id=sid,
                                       mode="login")

            # 2. 检查 admin 记录
            admin = Admin.query.filter_by(student_id=sid, name=name).first()
            if not admin:
                flash(
                    f"❌ {name}（学号 {sid}）尚未被加入管理员白名单，"
                    f"无法登录。请联系超管添加，或前往「📝 申请注册」。",
                    "warning",
                )
                return render_template("admin/login.html", name=name, student_id=sid,
                                       mode="login")

            # 3. 状态机
            if admin.is_pending_approval():
                flash(
                    f"⏳ {name}，您的注册申请正在审批中，请耐心等待超管审核。",
                    "info",
                )
                return render_template("admin/login.html", name=name, student_id=sid,
                                       mode="login")
            if admin.is_awaiting_activation():
                # 已批准但未设密码 → 完成注册
                if len(pw) < 6:
                    flash("首次激活：密码至少 6 位", "warning")
                    return render_template("admin/login.html", name=name, student_id=sid,
                                           mode="activate")
                admin.set_password(pw)
                admin.status = "active"
                admin.last_login_at = datetime.utcnow()
                db.session.commit()
                log_modification(name, None,
                                 f"激活账号：{name}（{'超管' if admin.is_super_admin else '普管'}）")
                db.session.commit()
                flash(f"🎉 账号激活成功，欢迎 {name}！", "success")
            else:
                # active → 验证密码
                if not admin.check_password(pw):
                    flash("密码错误", "danger")
                    return render_template("admin/login.html", name=name, student_id=sid,
                                           mode="login")

            # 4. 登录
            session.clear()
            session["admin_id"] = admin.id
            session["admin_name"] = admin.name
            session["is_super_admin"] = bool(admin.is_super_admin)
            session.permanent = remember
            admin.last_login_at = datetime.utcnow()
            db.session.commit()
            log_modification(admin.name, None,
                             f"管理员登录（{'超管' if admin.is_super_admin else '普管'}）")
            db.session.commit()
            if not admin.is_super_admin and Admin.query.filter_by(is_super_admin=True).count() == 0:
                admin.is_super_admin = True
                db.session.commit()
                session["is_super_admin"] = True
            flash(f"欢迎，{admin.name}！", "success")
            nxt = request.args.get("next") or url_for("admin_dashboard")
            return redirect(nxt)

        return render_template("admin/login.html", mode="login")

    # ---------- 管理员申请注册 ----------
    @app.route("/admin/apply", methods=["POST"])
    def admin_apply():
        """提交注册申请（无白名单的成员可走此路）
        - 必须在 members 表
        - 不可与现有 admin 重名+学号
        - 创建 status='pending' 记录
        """
        name = (request.form.get("name") or "").strip()
        sid = (request.form.get("student_id") or "").strip()
        msg = (request.form.get("application_message") or "").strip()
        if not name or not sid:
            flash("姓名和学号不能为空", "danger")
            return redirect(url_for("admin_login"))
        member = Member.query.filter_by(name=name, student_id=sid).first()
        if not member:
            flash(f"❌ {name}（学号 {sid}）不在社团成员名单中，无法申请", "danger")
            return redirect(url_for("admin_login"))
        existing = Admin.query.filter_by(student_id=sid, name=name).first()
        if existing:
            if existing.is_activated():
                flash(f"⚠️ {name} 您已是管理员，请直接登录", "warning")
            elif existing.is_awaiting_activation():
                flash(f"⏳ {name} 您的申请已批准，请前往登录页设置密码", "info")
            else:
                flash(f"⏳ {name} 您的申请已在审批中，请耐心等待", "info")
            return redirect(url_for("admin_login"))
        # 创建 pending
        a = Admin(
            name=name, student_id=sid,
            password_hash="",
            status="pending",
            application_message=msg or "（无）",
        )
        db.session.add(a)
        db.session.commit()
        log_modification(name, None, f"提交管理员注册申请：{name}（{msg[:40] if msg else '无说明'}）")
        db.session.commit()
        flash(
            f"📨 申请已提交！{name}，请等待超管审批。审批通过后您可在此页设置密码。",
            "success",
        )
        return redirect(url_for("admin_login"))

    @app.route("/admin/reset_password", methods=["POST"])
    def admin_reset_password():
        """忘记密码：通过姓名+学号+手机号+年级验证身份，重置为随机密码并显示"""
        name = (request.form.get("name") or "").strip()
        sid = (request.form.get("student_id") or "").strip()
        phone = (request.form.get("phone") or "").strip()
        grade = normalize_grade(request.form.get("grade"))
        if not name or not sid:
            flash("姓名和学号不能为空", "danger")
            return redirect(url_for("admin_login"))
        admin = Admin.query.filter_by(name=name, student_id=sid).first()
        if not admin:
            flash(f"未找到管理员「{name}（学号 {sid}）」，请检查输入", "danger")
            return redirect(url_for("admin_login"))
        if not admin.is_activated():
            flash(f"「{name}」尚未激活，请直接设置密码登录", "warning")
            return redirect(url_for("admin_login"))
        member = Member.query.filter_by(name=name, student_id=sid).first()
        if not member:
            flash(f"未找到社员信息，请联系管理员", "danger")
            return redirect(url_for("admin_login"))
        if phone and phone != (member.phone or ""):
            flash(f"手机号不匹配，请检查输入", "danger")
            return redirect(url_for("admin_login"))
        if grade and grade != (member.grade or ""):
            flash(f"年级不匹配，请检查输入", "danger")
            return redirect(url_for("admin_login"))
        import string as _str, random as _rnd
        new_pw = ''.join(_rnd.choices(_str.ascii_letters + _str.digits, k=10))
        admin.set_password(new_pw)
        db.session.commit()
        log_modification("系统", None, f"重置密码：{name}（{sid}）")
        db.session.commit()
        flash(
            f"✅ 密码已重置。<br><br>"
            f"👤 姓名：<b>{name}</b><br>"
            f"🎓 学号：<b>{sid}</b><br>"
            f"🔑 新密码：<b style='font-size:18px;'>{new_pw}</b><br><br>"
            f"请立即登录并修改密码 ⚠️",
            "success",
        )
        return redirect(url_for("admin_login"))

    @app.route("/admin/logout")
    @admin_required
    def admin_logout():
        name = session.get("admin_name", "")
        session.clear()
        log_modification(name, None, "管理员退出登录")
        db.session.commit()
        flash("已退出登录", "info")
        return redirect(url_for("index"))

    # ---------- 管理员白名单管理（超管） ----------
    @app.route("/admin/whitelist")
    @super_admin_required
    def admin_whitelist():
        pending_admins = Admin.query.filter_by(status="pending").order_by(Admin.created_at.desc()).all()
        admins = Admin.query.filter(Admin.status != "pending").order_by(
            Admin.is_root.desc(), Admin.is_super_admin.desc(), Admin.created_at
        ).all()
        members_list = [{"id": m.id, "name": m.name, "student_id": m.student_id}
                        for m in Member.query.order_by(Member.name).all()]
        return render_template("admin/whitelist.html",
                               admins=admins,
                               pending_admins=pending_admins,
                               super_count=sum(1 for a in admins if a.is_super_admin),
                               norm_count=sum(1 for a in admins if not a.is_super_admin),
                               active_count=sum(1 for a in admins if a.is_activated()),
                               pending_count=len(pending_admins),
                               perm_choices=ADMIN_PERMISSIONS,
                               ADMIN_PERMISSIONS=ADMIN_PERMISSIONS,
                               members_json=_json.dumps(members_list, ensure_ascii=False))

    @app.route("/admin/whitelist/add", methods=["POST"])
    @super_admin_required
    def admin_whitelist_add():
        """添加管理员（超管预添加成员）
        - 必填：姓名（从成员库自动补全学号）
        - 密码：可选（不填 = 待激活）
        - 角色/权限：均有默认值
        """
        import json as _json
        name = (request.form.get("name") or "").strip()
        sid = (request.form.get("student_id") or "").strip()
        pw = request.form.get("password", "")
        is_super = request.form.get("is_super_admin") == "1"
        perms = request.form.getlist("permissions")

        if not name:
            flash("姓名不能为空", "danger")
            return redirect(url_for("admin_whitelist"))
        # 学号未填时自动从成员库匹配
        if not sid:
            match = Member.query.filter_by(name=name).first()
            if match:
                sid = match.student_id
            else:
                flash(f"未找到成员「{name}」，请确认姓名或手动填写学号", "danger")
                return redirect(url_for("admin_whitelist"))
        else:
            # 验证学号与姓名匹配
            member = Member.query.filter_by(name=name, student_id=sid).first()
            if not member:
                flash(f"姓名「{name}」与学号「{sid}」不匹配", "danger")
                return redirect(url_for("admin_whitelist"))
        if pw and len(pw) < 6:
            flash("密码至少 6 位（留空表示待激活）", "danger")
            return redirect(url_for("admin_whitelist"))
        if Admin.query.filter_by(student_id=sid).first():
            flash(f"学号 {sid} 已在白名单中", "danger")
            return redirect(url_for("admin_whitelist"))

        # 超管不需要 permissions；普管默认 dashboard
        if is_super:
            perms_str = ""
        else:
            valid_keys = {k for k, _, _ in ADMIN_PERMISSIONS}
            perms = [p for p in perms if p in valid_keys]
            if not perms:
                perms = ["dashboard"]
            perms_str = _json.dumps(perms, ensure_ascii=False)

        a = Admin(
            name=name, student_id=sid,
            is_super_admin=is_super,
            permissions=perms_str,
            status="approved" if not pw else "active",
            approved_at=datetime.utcnow(),
            approved_by=session.get("admin_name", ""),
        )
        if pw:
            a.set_password(pw)
        db.session.add(a)
        db.session.commit()
        if is_super:
            perm_desc = "全权限"
        else:
            perm_labels = [n for k, n, _ in ADMIN_PERMISSIONS if k in perms]
            perm_desc = f"{len(perms)} 项权限：{'、'.join(perm_labels[:3])}{'…' if len(perms) > 3 else ''}"
        state = "active（已设密码）" if pw else "approved（待激活）"
        log_modification(session["admin_name"], None,
                         f"预添加管理员：{name}（{'超管' if is_super else '普管'}，{state}，{perm_desc}）")
        db.session.commit()
        flash(f"已添加管理员：{name}（{state}，{perm_desc}）", "success")
        return redirect(url_for("admin_whitelist"))

    # ---------- 审批申请 ----------
    @app.route("/admin/whitelist/<int:aid>/approve", methods=["POST"])
    @super_admin_required
    def admin_whitelist_approve(aid):
        """审批通过：pending → approved（无密码 = 待激活）
        超管可在此处为申请人设置初始密码（可选），或让申请人自行设置。
        """
        import json as _json
        a = Admin.query.get_or_404(aid)
        if a.status != "pending":
            flash(f"{a.name} 不在待审批状态（当前：{a.status}）", "warning")
            return redirect(url_for("admin_whitelist"))
        is_super = request.form.get("is_super_admin") == "1"
        pw = request.form.get("password", "")
        perms = request.form.getlist("permissions")
        if pw and len(pw) < 6:
            flash("密码至少 6 位（留空 = 让用户自行设置）", "danger")
            return redirect(url_for("admin_whitelist"))
        if is_super:
            a.is_super_admin = True
            a.permissions = ""
        else:
            a.is_super_admin = False
            valid_keys = {k for k, _, _ in ADMIN_PERMISSIONS}
            perms = [p for p in perms if p in valid_keys]
            if not perms:
                flash("请至少勾选一项权限（或设为超级管理员）", "warning")
                return redirect(url_for("admin_whitelist"))
            # 默认勾选 dashboard（让新管理员至少能进入首页）
            if "dashboard" not in perms:
                perms.append("dashboard")
            a.permissions = _json.dumps(perms, ensure_ascii=False)
        a.status = "approved" if not pw else "active"
        a.approved_at = datetime.utcnow()
        a.approved_by = session.get("admin_name", "")
        if pw:
            a.set_password(pw)
        db.session.commit()
        role = "超管" if a.is_super_admin else "普管"
        state = "active（已设密码）" if pw else "approved（待激活）"
        log_modification(session["admin_name"], None,
                         f"审批通过：{a.name} → {role}（{state}）")
        db.session.commit()
        flash(f"✅ 已批准 {a.name} 的申请（{role}，{state}）", "success")
        return redirect(url_for("admin_whitelist"))

    @app.route("/admin/whitelist/<int:aid>/reject", methods=["POST"])
    @super_admin_required
    def admin_whitelist_reject(aid):
        """拒绝申请：直接删除 pending 记录"""
        a = Admin.query.get_or_404(aid)
        if a.status != "pending":
            flash(f"{a.name} 不在待审批状态（当前：{a.status}）", "warning")
            return redirect(url_for("admin_whitelist"))
        name = a.name
        msg = a.application_message
        db.session.delete(a)
        db.session.commit()
        log_modification(session["admin_name"], None,
                         f"拒绝 {name} 的申请（{msg[:30] if msg else '无说明'}）")
        db.session.commit()
        flash(f"已拒绝 {name} 的申请", "warning")
        return redirect(url_for("admin_whitelist"))

    @app.route("/admin/whitelist/<int:aid>/permissions", methods=["POST"])
    @super_admin_required
    def admin_whitelist_set_permissions(aid):
        """更新某管理员的权限（超管自身不可降权限）"""
        import json as _json
        a = Admin.query.get_or_404(aid)
        if a.is_super_admin:
            flash("超管默认拥有全部权限，无需单独设置", "info")
            return redirect(url_for("admin_whitelist"))
        perms = request.form.getlist("permissions")
        valid_keys = {k for k, _, _ in ADMIN_PERMISSIONS}
        perms = [p for p in perms if p in valid_keys]
        if not perms:
            flash("请至少勾选一项权限", "warning")
            return redirect(url_for("admin_whitelist"))
        a.permissions = _json.dumps(perms, ensure_ascii=False)
        db.session.commit()
        log_modification(session["admin_name"], None,
                         f"更新 {a.name} 的权限：{len(perms)} 项")
        db.session.commit()
        perm_labels = [n for k, n, _ in ADMIN_PERMISSIONS if k in perms]
        flash(f"已更新 {a.name} 的权限：{'、'.join(perm_labels)}", "success")
        return redirect(url_for("admin_whitelist"))

    @app.route("/admin/whitelist/<int:aid>/delete", methods=["POST"])
    @super_admin_required
    def admin_whitelist_delete(aid):
        a = Admin.query.get_or_404(aid)
        me_id = session.get("admin_id")
        # 规则 1：创始人 (is_root) 只有本人可删
        if a.is_root and a.id != me_id:
            flash("创始人只能由本人删除", "danger")
            return redirect(url_for("admin_whitelist"))
        # 规则 2：非超管不能删别人
        me = Admin.query.get(me_id)
        if a.id != me_id and not me.is_super_admin:
            flash("只有超级管理员可以删除其他管理员", "danger")
            return redirect(url_for("admin_whitelist"))
        # 不能删最后一个超管
        if a.is_super_admin and not a.is_root:
            super_count = Admin.query.filter_by(is_super_admin=True).count()
            if super_count <= 1:
                flash("系统至少需要 1 位超级管理员", "danger")
                return redirect(url_for("admin_whitelist"))
        # 不能删自己（创始人也允许自删，单独路径 /admin/whitelist/self/delete）
        if a.id == me_id and not a.is_root:
            flash("如需删除自己的账号，请前往「我的资料」", "info")
            return redirect(url_for("admin_whitelist"))
        name = a.name
        role = "创始人" if a.is_root else ("超管" if a.is_super_admin else "普管")
        db.session.delete(a)
        db.session.commit()
        log_modification(session["admin_name"], None, f"删除管理员：{name}（{role}）")
        db.session.commit()
        flash(f"已删除管理员：{name}", "warning")
        return redirect(url_for("admin_whitelist"))

    # 创始人自删专属路径（创始人只能自己删自己）
    @app.route("/admin/whitelist/self/delete", methods=["POST"])
    @admin_required
    def admin_whitelist_self_delete():
        me = Admin.query.get_or_404(session.get("admin_id"))
        if not me.is_root:
            flash("该路径仅创始人可用", "danger")
            return redirect(url_for("admin_profile"))
        # 创始人删除自己会清空系统，需特别提示
        name = me.name
        db.session.delete(me)
        db.session.commit()
        session.clear()
        flash(f"创始人 {name} 已删除账号。系统已无任何超管，请立即注册新超管。", "warning")
        return redirect(url_for("index"))

    @app.route("/admin/whitelist/<int:aid>/toggle_super", methods=["POST"])
    @super_admin_required
    def admin_whitelist_toggle_super(aid):
        a = Admin.query.get_or_404(aid)
        me_id = session.get("admin_id")
        me = Admin.query.get(me_id)
        # 规则：创始人不允许被任何超管降级（即使本人除外）
        if a.is_root and a.id != me_id:
            flash("创始人身份不可被修改", "danger")
            return redirect(url_for("admin_whitelist"))
        # 非超管不能改别人
        if a.id != me_id and not me.is_super_admin:
            flash("只有超级管理员可以修改其他管理员", "danger")
            return redirect(url_for("admin_whitelist"))
        # 若要降级当前超管，确保至少留 1 个超管
        if a.is_super_admin:
            super_count = Admin.query.filter_by(is_super_admin=True).count()
            if super_count <= 1:
                flash("系统至少需要 1 位超级管理员", "danger")
                return redirect(url_for("admin_whitelist"))
        a.is_super_admin = not a.is_super_admin
        db.session.commit()
        log_modification(session["admin_name"], None,
                         f"{a.name} 已被{'设为' if a.is_super_admin else '取消'}超管")
        db.session.commit()
        flash(f"已将 {a.name} {'设为' if a.is_super_admin else '取消'}超管身份", "success")
        return redirect(url_for("admin_whitelist"))

    @app.route("/admin/whitelist/<int:aid>/reset_password", methods=["POST"])
    @super_admin_required
    def admin_whitelist_reset_password(aid):
        a = Admin.query.get_or_404(aid)
        new_pw = request.form.get("new_password", "")
        if len(new_pw) < 6:
            flash("新密码至少 6 位", "danger")
            return redirect(url_for("admin_whitelist"))
        a.set_password(new_pw)
        db.session.commit()
        log_modification(session["admin_name"], None, f"重置 {a.name} 的密码")
        db.session.commit()
        flash(f"已重置 {a.name} 的密码", "success")
        return redirect(url_for("admin_whitelist"))

    @app.route("/admin/whitelist/self/reset_password", methods=["POST"])
    @admin_required
    def admin_whitelist_self_reset_password():
        """任何管理员都可以重置自己的密码（如果他们记得旧密码）"""
        me = Admin.query.get_or_404(session.get("admin_id"))
        old_pw = request.form.get("old_password", "")
        new_pw = request.form.get("new_password", "")
        if not me.check_password(old_pw):
            flash("当前密码错误", "danger")
            return redirect(url_for("admin_profile"))
        if len(new_pw) < 6:
            flash("新密码至少 6 位", "danger")
            return redirect(url_for("admin_profile"))
        me.set_password(new_pw)
        db.session.commit()
        log_modification(me.name, None, "修改了自己的密码")
        db.session.commit()
        flash("密码已更新", "success")
        return redirect(url_for("admin_profile"))

    @app.route("/admin/profile", methods=["GET", "POST"])
    @admin_required
    def admin_profile():
        me = Admin.query.get_or_404(session["admin_id"])
        if request.method == "POST":
            old = request.form.get("old_password", "")
            new = request.form.get("new_password", "")
            new2 = request.form.get("new_password2", "")
            if not me.check_password(old):
                flash("当前密码错误", "danger")
            elif len(new) < 6:
                flash("新密码至少 6 位", "danger")
            elif new != new2:
                flash("两次输入的新密码不一致", "danger")
            else:
                me.set_password(new)
                db.session.commit()
                log_modification(me.name, None, "修改了自己的密码")
                db.session.commit()
                flash("密码已更新", "success")
                return redirect(url_for("admin_profile"))
        return render_template("admin/profile.html", me=me)

    # ---------- 管理员后台首页 ----------
    @app.route("/admin")
    @admin_required
    def admin_dashboard():
        a_act = db.session.query(func.sum(ActivityParticipant.a_hours)).scalar() or 0
        a_ir = db.session.query(func.sum(IrregularParticipant.a_hours)).scalar() or 0
        b_act = db.session.query(func.sum(ActivityParticipant.b_hours)).scalar() or 0
        b_ir = db.session.query(func.sum(IrregularParticipant.b_hours)).scalar() or 0
        a_hours = float(a_act) + float(a_ir)
        b_hours = float(b_act) + float(b_ir)
        stats = {
            "members": Member.query.count(),
            "blacklist": Blacklist.query.count(),
            "projects": Project.query.filter_by(active=True).count(),
            "activities": Activity.query.count(),
            "a_hours": round(a_hours, 1),
            "b_hours": round(b_hours, 1),
            "hours": round(a_hours + b_hours, 1),
            "points": round(sum(m.total_points() for m in Member.query.all()), 2),
            "queries": QueryLog.query.count(),
            "signoffs": SignoffLog.query.count(),
        }
        # 最近活动
        recent_activities = Activity.query.order_by(Activity.id.desc()).limit(8).all()
        # 最近签收
        recent_signoffs = SignoffLog.query.order_by(SignoffLog.id.desc()).limit(8).all()
        # 未签收社员
        unsigned = []
        for m in Member.query.all():
            if not m.signoffs:
                unsigned.append(m)
        return render_template(
            "admin/dashboard.html", stats=stats,
            recent_activities=recent_activities,
            recent_signoffs=recent_signoffs,
            unsigned_count=len(unsigned),
        )

    # ---------- 社员管理 ----------
    @app.route("/admin/members")
    @permission_required("members")
    def admin_members():
        keyword = request.args.get("q", "").strip()
        identity_filter = request.args.get("identity", "all").strip()
        q = Member.query
        if keyword:
            q = q.filter(
                db.or_(
                    Member.name.like(f"%{keyword}%"),
                    Member.student_id.like(f"%{keyword}%"),
                )
            )
        if identity_filter == "member":
            q = q.filter(Member.identity == "社员")
        elif identity_filter == "longterm":
            q = q.filter(Member.identity == "长期项目组成员")
        members = q.order_by(Member.id.asc()).all()
        return render_template(
            "admin/members.html",
            members=members, keyword=keyword,
            identity_filter=identity_filter,
            all_count=Member.query.count(),
            member_count=Member.query.filter_by(identity="社员").count(),
            longterm_count=Member.query.filter_by(identity="长期项目组成员").count(),
        )

    @app.route("/admin/members/new", methods=["GET", "POST"])
    @permission_required("members")
    def admin_member_new():
        if request.method == "POST":
            name = (request.form.get("name") or "").strip()
            if not name:
                flash("姓名不能为空", "danger")
                return redirect(url_for("admin_member_new"))
            lt_id = request.form.get("long_term_group_id", "").strip()
            m = Member(
                name=name,
                grade=normalize_grade(request.form.get("grade")),
                student_id=(request.form.get("student_id") or "").strip(),
                department=(request.form.get("department") or "").strip(),
                identity=(request.form.get("identity") or "社员").strip(),
                long_term_group_id=int(lt_id) if lt_id else None,
                phone=(request.form.get("phone") or "").strip(),
            )
            db.session.add(m)
            db.session.commit()
            log_modification(session.get("admin_name", "管理员"), m.id, f"新增成员：{m.name}")
            db.session.commit()
            flash(f"已添加成员【{m.name}】", "success")
            return redirect(url_for("admin_members"))
        return render_template(
            "admin/member_edit.html", member=None,
            departments=all_departments(),
            long_term_groups=LongTermGroup.query.filter_by(active=True).order_by(LongTermGroup.sort_order).all(),
        )

    @app.route("/admin/members/<int:mid>/edit", methods=["GET", "POST"])
    @permission_required("members")
    def admin_member_edit(mid):
        m = Member.query.get_or_404(mid)
        if request.method == "POST":
            changes = []
            for field in ["name", "grade", "student_id", "department", "phone", "identity"]:
                new = normalize_grade(request.form.get(field)) if field == "grade" else (request.form.get(field) or "").strip()
                old = getattr(m, field) or ""
                if new != old:
                    changes.append(f"字段【{field}】变更：\n  - 原值：{old}\n  - 新值：{new}")
                    setattr(m, field, new)
            # 长期项目组
            lt_id_str = request.form.get("long_term_group_id", "").strip()
            new_lt_id = int(lt_id_str) if lt_id_str else None
            old_lt_id = m.long_term_group_id
            if new_lt_id != old_lt_id:
                old_name = m.long_term_group.name if m.long_term_group else "（无）"
                new_name = LongTermGroup.query.get(new_lt_id).name if new_lt_id else "（无）"
                changes.append(f"字段【long_term_group_id】变更：\n  - 原值：{old_name}\n  - 新值：{new_name}")
                m.long_term_group_id = new_lt_id
            if changes:
                log_modification(session.get("admin_name", "管理员"), m.id, "\n\n".join(changes))
            db.session.commit()
            flash("保存成功", "success")
            return redirect(url_for("admin_members"))
        return render_template(
            "admin/member_edit.html", member=m,
            departments=all_departments(),
            long_term_groups=LongTermGroup.query.filter_by(active=True).order_by(LongTermGroup.sort_order).all(),
        )

    @app.route("/admin/members/<int:mid>/delete", methods=["POST"])
    @permission_required("members")
    def admin_member_delete(mid):
        m = Member.query.get_or_404(mid)
        name = m.name
        db.session.delete(m)
        log_modification(session.get("admin_name", "管理员"), None, f"删除社员：{name}")
        db.session.commit()
        flash(f"已删除社员【{name}】及其所有积分记录", "warning")
        return redirect(url_for("admin_members"))

    # ---------- 批量导入社员 ----------
    # 列名识别映射（支持中英文/同义词）
    def _parse_pasted_text(text):
        """解析用户从 Excel/WPS 复制粘贴的文本 → 二维 rows
        支持 Tab 分隔（Excel 复制）、逗号分隔（CSV 复制）、混合换行
        """
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            # 优先 Tab 分隔（Excel 默认），其次逗号
            if "\t" in line:
                cells = line.split("\t")
            elif "," in line:
                # 简单 CSV：不处理引号转义（够用）
                cells = [c.strip().strip('"').strip("'") for c in line.split(",")]
            else:
                cells = [line]
            rows.append([c.strip() for c in cells])
        return rows

    def _looks_like_header(row):
        """判断第一行是否像表头（而非数据行）。
        启发式：若含 ≥2 个明确字段名关键词 → 表头；
        若行里出现学号/手机号/日期/时间等典型数据 → 数据行
        """
        if not row:
            return False
        header_keywords = {
            "姓名", "名字", "name", "fullname",
            "年级", "grade", "level", "入学年份", "届数",
            "学号", "sid", "student_id", "studentid", "编号",
            "部门", "department", "dept", "所属部门", "组别",
            "身份", "identity", "角色", "类型", "member_type",
            "长期项目组", "long_term_group", "项目组", "工作组", "lt_group",
            "手机号", "phone", "tel", "mobile", "联系方式", "电话",
        }
        data_patterns = [
            lambda s: s.isdigit() and len(s) >= 6,  # 长数字：学号/手机号
            lambda s: bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}$", s)),  # 日期
            lambda s: bool(re.match(r"^\d{1,2}:\d{2}", s)),  # 时间
            lambda s: bool(re.match(r"^\d+\.?\d*$", s)),  # 数字（含小数）
        ]
        header_hits = 0
        data_hits = 0
        for c in row:
            if c is None:
                continue
            s = str(c).strip()
            if not s:
                continue
            s_lower = s.lower()
            if s in header_keywords or s_lower in header_keywords:
                header_hits += 1
                continue
            for pat in data_patterns:
                if pat(s):
                    data_hits += 1
                    break
        # 规则：
        # - 至少 2 个字段名关键词 + 0 个数据模式 → 表头
        # - 任一数据模式命中 + 0 字段名 → 数据行
        # - 都有 → 取多者
        if header_hits >= 2 and data_hits == 0:
            return True
        if data_hits > 0 and header_hits == 0:
            return False
        if data_hits >= 2 and header_hits < data_hits:
            return False
        if header_hits >= 2:
            return True
        return False

    COLUMN_ALIASES = {
        "name": ["姓名", "名字", "name", "fullname", "member_name", "社员姓名", "成员姓名"],
        "grade": ["年级", "grade", "level", "入学年份", "届数"],
        "student_id": ["学号", "student_id", "studentid", "sid", "编号"],
        "department": ["部门", "department", "dept", "所属部门", "组别"],
        "identity": ["身份", "identity", "角色", "类型", "member_type"],
        "long_term_group": ["长期项目组", "long_term_group", "项目组", "工作组", "lt_group"],
        "phone": ["手机号", "phone", "tel", "mobile", "联系方式", "电话"],
    }

    def detect_columns(headers):
        """根据表头自动识别列对应。返回 {字段名: 列索引} 字典。"""
        mapping = {}
        used = set()
        for field, aliases in COLUMN_ALIASES.items():
            for i, h in enumerate(headers):
                if i in used:
                    continue
                h_clean = str(h or "").strip().lower().replace(" ", "")
                if h_clean in [a.lower().replace(" ", "") for a in aliases]:
                    mapping[field] = i
                    used.add(i)
                    break
        return mapping

    @app.route("/admin/members/import", methods=["GET", "POST"])
    @permission_required("members")
    def admin_member_import():
        if request.method == "GET":
            return render_template("admin/member_import.html",
                                   preview=None, error=None, departments=all_departments())

        action = request.form.get("action", "preview")

        if action == "download_template":
            return _download_member_template()

        # 确认导入：从 session 读取之前解析的 rows
        if action == "confirm":
            cached = session.get("import_data")
            if not cached:
                return render_template("admin/member_import.html",
                                       preview=None, error="会话已过期，请重新上传文件",
                                       departments=all_departments())
            try:
                col_name = int(request.form.get("col_name", -1))
                col_grade = int(request.form.get("col_grade", -1))
                col_sid = int(request.form.get("col_sid", -1))
                col_dept = int(request.form.get("col_dept", -1))
                col_phone = int(request.form.get("col_phone", -1))
                col_identity = int(request.form.get("col_identity", -1))
                col_ltgroup = int(request.form.get("col_ltgroup", -1))
            except (ValueError, TypeError):
                return render_template("admin/member_import.html",
                                       preview=None, error="列对应配置错误",
                                       departments=all_departments())

            if col_name < 0:
                return render_template("admin/member_import.html",
                                       preview=None, error="必须指定「姓名」所在的列",
                                       departments=all_departments())

            default_dept = request.form.get("default_department", "").strip()
            rows = cached["rows"]

            success, skipped, failed = 0, 0, 0
            failed_details = []
            existing_sids = {m.student_id for m in Member.query.all() if m.student_id}
            existing_names = {m.name for m in Member.query.all()}

            for idx, r in enumerate(rows, 1):
                if not any(r):
                    continue
                try:
                    name = str(r[col_name]).strip() if col_name < len(r) and r[col_name] is not None else ""
                except IndexError:
                    name = ""
                if not name:
                    failed += 1
                    failed_details.append(f"第{idx}行：姓名为空")
                    continue

                if name in existing_names:
                    skipped += 1
                    continue

                def get_col(idx2):
                    if idx2 < 0 or idx2 >= len(r):
                        return ""
                    v = r[idx2]
                    return str(v).strip() if v is not None else ""

                sid = get_col(col_sid)
                if sid and sid in existing_sids:
                    skipped += 1
                    continue

                dept = get_col(col_dept) or default_dept
                # 身份
                identity_raw = get_col(col_identity) or "社员"
                if "长期" in identity_raw or "long" in identity_raw.lower():
                    identity_val = "长期项目组成员"
                else:
                    identity_val = "社员"
                # 长期项目组（按名字匹配）
                lt_name = get_col(col_ltgroup)
                lt_id = None
                if lt_name:
                    g = LongTermGroup.query.filter_by(name=lt_name).first()
                    lt_id = g.id if g else None
                m = Member(
                    name=name,
                    grade=normalize_grade(get_col(col_grade)),
                    student_id=sid,
                    department=dept,
                    identity=identity_val,
                    long_term_group_id=lt_id,
                    phone=get_col(col_phone),
                )
                db.session.add(m)
                existing_names.add(name)
                if sid:
                    existing_sids.add(sid)
                success += 1

            if success > 0:
                log_modification(session.get("admin_name", "管理员"), None,
                                 f"批量导入社员：成功 {success} 条，跳过 {skipped} 条，失败 {failed} 条")
            db.session.commit()
            session.pop("import_data", None)

            flash(f"导入完成：✅ 成功 {success} 条，跳过重复 {skipped} 条，失败 {failed} 条", "success")
            if failed_details:
                flash("失败明细：" + "；".join(failed_details[:10]), "warning")
            return redirect(url_for("admin_members"))

        # action == preview: 解析上传文件 / 粘贴文本
        paste_text = request.form.get("paste_text", "").strip()
        if paste_text:
            rows = _parse_pasted_text(paste_text)
        else:
            f = request.files.get("file")
            if not f or not f.filename:
                return render_template("admin/member_import.html",
                                       preview=None, error="请选择文件或粘贴数据", departments=all_departments())

            filename = f.filename.lower()
            try:
                if filename.endswith(".csv"):
                    content = f.read().decode("utf-8-sig")
                    reader = csv.reader(io.StringIO(content))
                    rows = list(reader)
                elif filename.endswith((".xlsx", ".xls")):
                    wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True, read_only=True)
                    # 优先找含数据最多的 sheet（避免选到"说明"sheet）
                    best_ws = None
                    best_count = 0
                    for ws in wb.worksheets:
                        ws_rows = [list(r) for r in ws.iter_rows(values_only=True)]
                        non_empty = sum(1 for r in ws_rows if any(c is not None and str(c).strip() for c in r))
                        if non_empty > best_count:
                            best_count = non_empty
                            best_ws = ws
                            rows = ws_rows
                    wb.close()
                    if best_ws is None:
                        return render_template("admin/member_import.html",
                                               preview=None, error="文件中没有数据",
                                               departments=all_departments())
                else:
                    return render_template("admin/member_import.html",
                                           preview=None, error="仅支持 .xlsx / .xls / .csv 文件",
                                           departments=all_departments())
            except Exception as e:
                return render_template("admin/member_import.html",
                                       preview=None, error=f"文件解析失败：{e}",
                                       departments=all_departments())

        if not rows:
            return render_template("admin/member_import.html",
                                   preview=None, error="文件为空", departments=all_departments())

        # 检测表头是否为空（无表头文件 → 全部当作数据）
        first_row = rows[0]
        first_row_has_text = any(c is not None and str(c).strip() for c in first_row)

        # 自动检测：若第一行看起来不像表头（如全是数字、或首列不是常见字段名），则当作数据行
        likely_header = first_row_has_text and _looks_like_header(first_row)

        if likely_header:
            headers = [str(c or "").strip() for c in first_row]
            data_rows = [r for r in rows[1:] if any(c is not None and str(c).strip() for c in r)]
        else:
            # 无表头：用列号占位（用户可手动选择）
            ncols = max(len(r) for r in rows)
            headers = [f"列{ i+1 }" for i in range(ncols)]
            data_rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]

        auto_mapping = detect_columns(headers)

        # 存到 session（转成可序列化的 list）
        session["import_data"] = {
            "headers": headers,
            "rows": [[("" if v is None else str(v).strip()) for v in r] for r in data_rows],
        }

        preview_rows = [list(r) for r in data_rows[:20]]
        return render_template("admin/member_import.html",
                               preview={
                                   "headers": headers,
                                   "rows": preview_rows,
                                   "total": len(data_rows),
                                   "mapping": auto_mapping,
                               },
                               error=None, departments=all_departments())

    def _download_member_template():
        """生成并返回成员导入模板 xlsx"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "成员名单"
        headers = ["姓名", "年级", "学号", "部门", "身份", "长期项目组", "手机号"]
        # 写入表头（带样式）
        header_fill = PatternFill("solid", start_color="4472C4")
        header_font = Font(bold=True, color="FFFFFF")
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = header_fill
            c.font = header_font
        # 写入示例
        sample = [
            ["张三", "2024级", "22301010001", "人资部", "社员", "", "13800000001"],
            ["李四", "2023级", "22202020002", "宣传部", "长期项目组成员", "心暖陇原", "13800000002"],
            ["王五", "2025级", "22503030003", "活动部", "社员", "心暖苗疆", "13800000003"],
        ]
        for r, row in enumerate(sample, 2):
            for i, v in enumerate(row, 1):
                ws.cell(row=r, column=i, value=v)
        # 备注行
        ws.cell(row=6, column=1, value="说明：")
        ws.cell(row=7, column=1, value="1. 姓名为必填项，其余可空")
        ws.cell(row=8, column=1, value="2. 列名支持中英文，系统会自动识别")
        ws.cell(row=9, column=1, value="3. 部门可填写：人资部/综管部/文体部/活动部/宣传部/创设部/外联部/财务部")
        ws.cell(row=10, column=1, value="4. 身份：「社员」或「长期项目组成员」（含「长期」字样自动归入后者）")
        ws.cell(row=11, column=1, value="5. 长期项目组：按组名匹配（须在「长期项目组」模块先创建）")
        # 列宽
        for i, w in enumerate([12, 12, 16, 12, 16, 18, 14], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf, as_attachment=True,
            download_name="爱心社成员导入模板.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------- 批量删除社员 ----------
    @app.route("/admin/members/batch_delete", methods=["POST"])
    @permission_required("members")
    def admin_member_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请先选择要删除的社员", "warning")
            return redirect(url_for("admin_members"))
        try:
            ids = [int(x) for x in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_members"))

        members = Member.query.filter(Member.id.in_(ids)).all()
        names = [m.name for m in members]
        for m in members:
            db.session.delete(m)
        log_modification(session.get("admin_name", "管理员"), None,
                         f"批量删除社员（{len(members)} 人）：{'、'.join(names[:20])}"
                         + (" 等" if len(names) > 20 else ""))
        db.session.commit()
        flash(f"已批量删除 {len(members)} 名社员及其所有积分记录", "warning")
        return redirect(url_for("admin_members"))

    # ---------- 项目管理 ----------
    @app.route("/admin/projects")
    @permission_required("projects")
    def admin_projects():
        projects = Project.query.order_by(Project.sort_order).all()
        return render_template("admin/projects.html", projects=projects)

    @app.route("/admin/projects/new", methods=["POST"])
    @permission_required("projects")
    def admin_project_new():
        name = (request.form.get("name") or "").strip()
        if not name:
            flash("项目名不能为空", "danger")
            return redirect(url_for("admin_projects"))
        p = Project(
            name=name,
            description=(request.form.get("description") or "").strip(),
            sort_order=int(request.form.get("sort_order") or 0),
        )
        db.session.add(p)
        db.session.commit()
        flash(f"已添加项目【{p.name}】", "success")
        return redirect(url_for("admin_projects"))

    @app.route("/admin/projects/<int:pid>/edit", methods=["POST"])
    @permission_required("projects")
    def admin_project_edit(pid):
        p = Project.query.get_or_404(pid)
        for field in ["name", "description"]:
            v = (request.form.get(field) or "").strip()
            setattr(p, field, v)
        p.sort_order = int(request.form.get("sort_order") or 0)
        p.active = request.form.get("active") == "on"
        db.session.commit()
        flash("项目已更新", "success")
        return redirect(url_for("admin_projects"))

    @app.route("/admin/projects/<int:pid>/delete", methods=["POST"])
    @permission_required("projects")
    def admin_project_delete(pid):
        p = Project.query.get_or_404(pid)
        if p.activities:
            flash(f"项目【{p.name}】下还有 {len(p.activities)} 场活动，请先删除活动", "danger")
            return redirect(url_for("admin_projects"))
        db.session.delete(p)
        db.session.commit()
        flash(f"已删除项目【{p.name}】", "warning")
        return redirect(url_for("admin_projects"))

    @app.route("/admin/projects/batch_delete", methods=["POST"])
    @permission_required("projects")
    def admin_project_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请先选择要删除的项目", "warning")
            return redirect(url_for("admin_projects"))
        try:
            ids = [int(x) for x in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_projects"))
        projs = Project.query.filter(Project.id.in_(ids)).all()
        skipped, deleted = 0, 0
        names = []
        for p in projs:
            if p.activities:
                skipped += 1
                continue
            names.append(p.name)
            db.session.delete(p)
            deleted += 1
        db.session.commit()
        if deleted:
            log_modification(session.get("admin_name", "管理员"), None,
                             f"批量删除项目：{', '.join(names)}")
            db.session.commit()
        msg = f"已删除 {deleted} 个项目"
        if skipped:
            msg += f"，跳过 {skipped} 个（其下还有活动）"
        flash(msg, "success" if deleted else "warning")
        return redirect(url_for("admin_projects"))

    @app.route("/admin/projects/batch_active", methods=["POST"])
    @permission_required("projects")
    def admin_project_batch_active():
        ids = request.form.getlist("ids")
        active = request.form.get("active") == "1"
        if not ids:
            flash("请先选择项目", "warning")
            return redirect(url_for("admin_projects"))
        try:
            ids = [int(x) for x in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_projects"))
        n = Project.query.filter(Project.id.in_(ids)).update({"active": active}, synchronize_session=False)
        db.session.commit()
        log_modification(session.get("admin_name", "管理员"), None,
                         f"批量{'启用' if active else '禁用'}项目 {n} 个")
        db.session.commit()
        flash(f"已{'启用' if active else '禁用'} {n} 个项目", "success")
        return redirect(url_for("admin_projects"))

    # ---------- 活动管理 ----------
    @app.route("/admin/activities")
    @permission_required("activities")
    def admin_activities():
        pid = request.args.get("project_id", type=int)
        q = Activity.query
        if pid:
            q = q.filter_by(project_id=pid)
        activities = q.order_by(Activity.id.desc()).all()
        projects = Project.query.order_by(Project.sort_order).all()
        totals = {
            "served": sum(a.served_count or 0 for a in activities),
            "a_hours": sum(a.total_a_hours() for a in activities),
            "b_hours": sum(a.total_b_hours() for a in activities),
            "participants": sum(a.participant_count() for a in activities),
        }
        return render_template("admin/activities.html", activities=activities,
                               projects=projects, current_pid=pid, totals=totals)

    @app.route("/admin/activities/new", methods=["GET", "POST"])
    @permission_required("activities")
    def admin_activity_new():
        if request.method == "POST":
            pid = request.form.get("project_id", type=int)
            if not pid:
                flash("请选择项目", "danger")
                return redirect(url_for("admin_activity_new"))
            a = Activity(
                project_id=pid,
                sequence=int(request.form.get("sequence") or 0),
                date=(request.form.get("date") or "").strip(),
                time=(request.form.get("time") or "").strip(),
                location=(request.form.get("location") or "").strip(),
                content=(request.form.get("content") or "").strip(),
                served_count=int(request.form.get("served_count") or 0),
                a_hours=float(request.form.get("a_hours") or 0),
                b_hours=float(request.form.get("b_hours") or 0),
                semester=(request.form.get("semester") or current_semester()).strip(),
                created_by=session.get("admin_name", "管理员"),
            )
            db.session.add(a)
            db.session.commit()
            # 批量录入参与者
            try:
                pdata = json.loads(request.form.get("participants_data") or "[]")
            except Exception:
                pdata = []
            added = 0
            for item in pdata:
                mid = item.get("member_id")
                if not mid:
                    continue
                a_h = float(item.get("a_hours") or 0)
                b_h = float(item.get("b_hours") or 0)
                if a_h <= 0 and b_h <= 0:
                    continue
                if Member.query.get(mid) is None:
                    continue
                db.session.add(ActivityParticipant(
                    activity_id=a.id, member_id=mid,
                    a_hours=a_h, b_hours=b_h,
                ))
                added += 1
            db.session.commit()
            flash(f"已添加活动：{a.content or a.date}（{added} 位参与者）", "success")
            return redirect(url_for("admin_activity_detail", aid=a.id))
        projects = Project.query.filter_by(active=True).order_by(Project.sort_order).all()
        members = Member.query.order_by(Member.name).all()
        return render_template("admin/activity_edit.html", activity=None,
                               projects=projects, members=members, participants=[],
                               departments=all_departments())

    @app.route("/admin/activities/<int:aid>")
    @permission_required("activities")
    def admin_activity_detail(aid):
        a = Activity.query.get_or_404(aid)
        members = Member.query.order_by(Member.name).all()
        return render_template("admin/activity_detail.html", activity=a, members=members)

    @app.route("/admin/activities/<int:aid>/edit", methods=["GET", "POST"])
    @permission_required("activities")
    def admin_activity_edit(aid):
        a = Activity.query.get_or_404(aid)
        if request.method == "POST":
            for field in ["date", "time", "location", "content", "semester"]:
                a.__setattr__(field, (request.form.get(field) or "").strip())
            a.sequence = int(request.form.get("sequence") or 0)
            a.served_count = int(request.form.get("served_count") or 0)
            a.a_hours = float(request.form.get("a_hours") or 0)
            a.b_hours = float(request.form.get("b_hours") or 0)
            pid = request.form.get("project_id", type=int)
            if pid:
                a.project_id = pid
            # 同步参与者
            try:
                pdata = json.loads(request.form.get("participants_data") or "[]")
            except Exception:
                pdata = []
            existing = {ap.member_id: ap for ap in a.participants}
            sent_ids = set()
            for item in pdata:
                mid = item.get("member_id")
                if not mid:
                    continue
                a_h = float(item.get("a_hours") or 0)
                b_h = float(item.get("b_hours") or 0)
                if a_h <= 0 and b_h <= 0:
                    continue
                if Member.query.get(mid) is None:
                    continue
                sent_ids.add(int(mid))
                if int(mid) in existing:
                    existing[int(mid)].a_hours = a_h
                    existing[int(mid)].b_hours = b_h
                else:
                    db.session.add(ActivityParticipant(
                        activity_id=a.id, member_id=int(mid),
                        a_hours=a_h, b_hours=b_h,
                    ))
            for mid, ap in existing.items():
                if mid not in sent_ids:
                    db.session.delete(ap)
            db.session.commit()
            flash("活动信息及参与者已更新", "success")
            return redirect(url_for("admin_activity_detail", aid=a.id))
        projects = Project.query.filter_by(active=True).order_by(Project.sort_order).all()
        members = Member.query.order_by(Member.name).all()
        participants = a.participants
        return render_template("admin/activity_edit.html", activity=a, projects=projects,
                               members=members, participants=participants,
                               departments=all_departments())

    @app.route("/admin/activities/<int:aid>/delete", methods=["POST"])
    @permission_required("activities")
    def admin_activity_delete(aid):
        a = Activity.query.get_or_404(aid)
        db.session.delete(a)
        db.session.commit()
        flash("已删除活动", "warning")
        return redirect(url_for("admin_activities"))

    @app.route("/admin/activities/batch_delete", methods=["POST"])
    @permission_required("activities")
    def admin_activity_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请先选择要删除的活动", "warning")
            return redirect(url_for("admin_activities"))
        try:
            ids = [int(x) for x in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_activities"))
        n = Activity.query.filter(Activity.id.in_(ids)).delete(synchronize_session=False)
        db.session.commit()
        log_modification(session.get("admin_name", "管理员"), None,
                         f"批量删除活动 {n} 场")
        db.session.commit()
        flash(f"已删除 {n} 场活动", "success")
        return redirect(url_for("admin_activities"))

    @app.route("/admin/activities/<int:aid>/add_participant", methods=["POST"])
    @permission_required("activities")
    def admin_activity_add_participant(aid):
        a = Activity.query.get_or_404(aid)
        mid = request.form.get("member_id", type=int)
        a_h = float(request.form.get("a_hours") or 0)
        b_h = float(request.form.get("b_hours") or 0)
        if not mid or (a_h <= 0 and b_h <= 0):
            flash("请选择社员并填写 A 类或 B 类小时数", "danger")
            return redirect(url_for("admin_activity_detail", aid=aid))
        existing = ActivityParticipant.query.filter_by(activity_id=aid, member_id=mid).first()
        if existing:
            existing.a_hours = a_h
            existing.b_hours = b_h
        else:
            db.session.add(ActivityParticipant(
                activity_id=aid, member_id=mid,
                a_hours=a_h, b_hours=b_h,
            ))
        db.session.commit()
        flash("已添加参与者", "success")
        return redirect(url_for("admin_activity_detail", aid=aid))

    @app.route("/admin/activities/<int:aid>/import_participants", methods=["POST"])
    @permission_required("activities")
    def admin_activity_import_participants(aid):
        """通过文件批量导入活动参与者"""
        a = Activity.query.get_or_404(aid)
        file = request.files.get("file")
        if not file or not file.filename:
            flash("请上传文件", "danger")
            return redirect(url_for("admin_activity_detail", aid=aid))
        rows, err = _parse_upload_file(file)
        if err:
            flash(f"文件解析失败：{err}", "danger")
            return redirect(url_for("admin_activity_detail", aid=aid))
        cols = _detect_columns(rows[0], ACTIVITY_COL_ALIASES) if rows else {}
        col_name = cols.get("name", -1)
        col_sid = cols.get("student_id", -1)
        col_a = cols.get("a_hours", -1)
        col_b = cols.get("b_hours", -1)
        col_h = cols.get("service_hours", -1)
        if col_name < 0 and col_sid < 0:
            flash("文件中未找到「姓名」或「学号」列", "danger")
            return redirect(url_for("admin_activity_detail", aid=aid))
        added, updated, skipped = 0, 0, 0
        members_cache = {m.student_id: m for m in Member.query.all()}
        for row in rows[1:]:
            raw_name = str(row[col_name]).strip() if col_name >= 0 and col_name < len(row) else ""
            raw_sid = str(row[col_sid]).strip() if col_sid >= 0 and col_sid < len(row) else ""
            m = members_cache.get(raw_sid) or (Member.query.filter_by(name=raw_name).first() if raw_name else None)
            if not m:
                skipped += 1
                continue
            a_h = float(row[col_a]) if col_a >= 0 and col_a < len(row) and row[col_a] else 0
            b_h = float(row[col_b]) if col_b >= 0 and col_b < len(row) and row[col_b] else 0
            if col_h >= 0 and col_h < len(row) and row[col_h] and a_h <= 0 and b_h <= 0:
                v = float(row[col_h])
                if cols.get("a_hours", -1) < 0 <= cols.get("b_hours", -1):
                    b_h = v
                elif cols.get("b_hours", -1) < 0 <= cols.get("a_hours", -1):
                    a_h = v
                else:
                    a_h = v
            if a_h <= 0 and b_h <= 0:
                skipped += 1
                continue
            existing = ActivityParticipant.query.filter_by(activity_id=aid, member_id=m.id).first()
            if existing:
                existing.a_hours = a_h
                existing.b_hours = b_h
                updated += 1
            else:
                db.session.add(ActivityParticipant(activity_id=aid, member_id=m.id, a_hours=a_h, b_hours=b_h))
                added += 1
        db.session.commit()
        flash(f"✅ 导入完成：新增 {added} 人，更新 {updated} 人，跳过 {skipped} 条", "success")
        return redirect(url_for("admin_activity_detail", aid=aid))

    # ---------- 活动批量导入 ----------
    ACTIVITY_COL_ALIASES = {
        "project": ["项目", "项目名", "project", "项目名称", "归属项目"],
        "date": ["日期", "date", "活动日期"],
        "time": ["时间", "time", "活动时间", "时段"],
        "location": ["地点", "location", "活动地点"],
        "content": ["活动内容", "内容", "content", "主题"],
        "served_count": ["服务对象", "服务对象人数", "served", "served_count", "受益人数"],
        "name": ["姓名", "name", "志愿者", "志愿者姓名", "同学"],
        "student_id": ["学号", "sid", "student_id"],
        "a_hours": ["A类小时数", "A类小时", "A小时数", "A类(h)", "A(h)", "A时长"],
        "b_hours": ["B类小时数", "B类小时", "B小时数", "B类(h)", "B(h)", "B时长"],
        "service_hours": ["时长", "服务时长", "小时", "hours", "service_hours", "h"],
        "semester": ["学期", "semester"],
    }

    @app.route("/admin/activities/import", methods=["GET", "POST"])
    @permission_required("activities")
    def admin_activity_import():
        if request.method == "GET":
            return render_template("admin/activity_import.html",
                                   preview=None, error=None, projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

        action = request.form.get("action", "preview")

        if action == "download_template":
            return _download_activity_template()

        if action == "confirm":
            cached = session.get("activity_import_data")
            if not cached:
                return render_template("admin/activity_import.html",
                                       preview=None, error="会话已过期，请重新上传",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())
            rows = cached["rows"]
            try:
                col_project = int(request.form.get("col_project", -1))
                col_date = int(request.form.get("col_date", -1))
                col_time = int(request.form.get("col_time", -1))
                col_location = int(request.form.get("col_location", -1))
                col_content = int(request.form.get("col_content", -1))
                col_served = int(request.form.get("col_served", -1))
                col_name = int(request.form.get("col_name", -1))
                col_sid = int(request.form.get("col_sid", -1))
                col_a_hours = int(request.form.get("col_a_hours", -1))
                col_b_hours = int(request.form.get("col_b_hours", -1))
                col_hours = int(request.form.get("col_hours", -1))  # 兼容老模板「时长」单列
                col_semester = int(request.form.get("col_semester", -1))
                default_project_id = request.form.get("default_project_id", type=int)
            except (ValueError, TypeError):
                return render_template("admin/activity_import.html",
                                       preview=None, error="列对应配置错误",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

            if col_name < 0:
                return render_template("admin/activity_import.html",
                                       preview=None, error="必须指定「姓名」列",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())
            if col_a_hours < 0 and col_b_hours < 0 and col_hours < 0:
                return render_template("admin/activity_import.html",
                                       preview=None, error="必须指定「A类小时数」「B类小时数」或「时长」列",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())
            if col_project < 0 and not default_project_id:
                return render_template("admin/activity_import.html",
                                       preview=None, error="必须指定「项目」列或默认项目",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

            def get_col(r, idx):
                if idx < 0 or idx >= len(r):
                    return ""
                v = r[idx]
                return str(v).strip() if v is not None else ""

            members_by_sid = {m.student_id: m for m in Member.query.all() if m.student_id}
            members_by_name = {m.name: m for m in Member.query.all()}
            projects_by_name = {p.name: p for p in Project.query.all()}

            created_activities = 0
            added_participants = 0
            failed = 0
            failed_details = []

            for idx, r in enumerate(rows, 1):
                try:
                    name = get_col(r, col_name)
                    if not name:
                        failed += 1
                        failed_details.append(f"第{idx}行：姓名为空")
                        continue
                    a_h = float(get_col(r, col_a_hours) or 0) if col_a_hours >= 0 else 0
                    b_h = float(get_col(r, col_b_hours) or 0) if col_b_hours >= 0 else 0
                    if a_h <= 0 and b_h <= 0:
                        # 兼容老模板：单一"时长"列 → 默认归入 A 类
                        if col_hours >= 0:
                            legacy = float(get_col(r, col_hours) or 0)
                            if legacy <= 0:
                                failed += 1
                                failed_details.append(f"第{idx}行：时长无效")
                                continue
                            a_h, b_h = legacy, 0
                        else:
                            failed += 1
                            failed_details.append(f"第{idx}行：A/B 小时数都为空")
                            continue

                    # 解析项目
                    pname = get_col(r, col_project)
                    if pname and pname in projects_by_name:
                        proj = projects_by_name[pname]
                    elif default_project_id:
                        proj = Project.query.get(default_project_id)
                    else:
                        failed += 1
                        failed_details.append(f"第{idx}行：项目未匹配【{pname}】")
                        continue

                    # 创建或匹配活动（按日期+地点+内容+项目匹配）
                    date_v = get_col(r, col_date) or datetime.now().strftime("%Y-%m-%d")
                    time_v = get_col(r, col_time)
                    loc_v = get_col(r, col_location)
                    content_v = get_col(r, col_content)
                    served_v = int(get_col(r, col_served) or 0)
                    sem_v = get_col(r, col_semester) or current_semester()

                    act = Activity.query.filter_by(
                        project_id=proj.id,
                        date=date_v, location=loc_v, content=content_v,
                    ).first()
                    if not act:
                        act = Activity(
                            project_id=proj.id,
                            date=date_v, time=time_v, location=loc_v,
                            content=content_v, served_count=served_v,
                            a_hours=0, b_hours=0,
                            semester=sem_v, created_by=session.get("admin_name", "管理员"),
                        )
                        db.session.add(act)
                        db.session.flush()
                        created_activities += 1

                    # 匹配社员
                    sid = get_col(r, col_sid)
                    if sid and sid in members_by_sid:
                        member = members_by_sid[sid]
                    elif name in members_by_name:
                        member = members_by_name[name]
                    else:
                        failed += 1
                        failed_details.append(f"第{idx}行：未找到社员【{name} / {sid}】")
                        continue

                    # 添加/更新参与
                    existing = ActivityParticipant.query.filter_by(
                        activity_id=act.id, member_id=member.id).first()
                    if existing:
                        existing.a_hours = (existing.a_hours or 0) + a_h
                        existing.b_hours = (existing.b_hours or 0) + b_h
                    else:
                        db.session.add(ActivityParticipant(
                            activity_id=act.id, member_id=member.id,
                            a_hours=a_h, b_hours=b_h,
                        ))
                    added_participants += 1
                except Exception as e:
                    failed += 1
                    failed_details.append(f"第{idx}行：{e}")

            log_modification(session.get("admin_name", "管理员"), None,
                             f"批量导入活动：创建 {created_activities} 场，新增 {added_participants} 条参与，失败 {failed} 条")
            db.session.commit()
            session.pop("activity_import_data", None)

            flash(f"导入完成：✅ 创建 {created_activities} 场活动，新增 {added_participants} 条参与记录，失败 {failed} 条", "success")
            if failed_details:
                flash("失败明细（前 10 条）：" + "；".join(failed_details[:10]), "warning")
            return redirect(url_for("admin_activities"))

        # preview
        f = request.files.get("file")
        if not f or not f.filename:
            return render_template("admin/activity_import.html",
                                   preview=None, error="请选择文件",
                                   projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())
        filename = f.filename.lower()
        try:
            if filename.endswith(".csv"):
                content = f.read().decode("utf-8-sig")
                reader = csv.reader(io.StringIO(content))
                rows = list(reader)
            elif filename.endswith((".xlsx", ".xls")):
                wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True, read_only=True)
                ws = wb.active
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                wb.close()
            else:
                return render_template("admin/activity_import.html",
                                       preview=None, error="仅支持 .xlsx/.xls/.csv",
                                       projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())
        except Exception as e:
            return render_template("admin/activity_import.html",
                                   preview=None, error=f"文件解析失败：{e}",
                                   projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

        if not rows:
            return render_template("admin/activity_import.html",
                                   preview=None, error="文件为空",
                                   projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

        headers = [str(c or "").strip() for c in rows[0]]
        # 自动识别
        mapping = {}
        used = set()
        for field, aliases in ACTIVITY_COL_ALIASES.items():
            for i, h in enumerate(headers):
                if i in used:
                    continue
                h_clean = str(h or "").strip().lower().replace(" ", "")
                if h_clean in [a.lower().replace(" ", "") for a in aliases]:
                    mapping[field] = i
                    used.add(i)
                    break
        data_rows = [r for r in rows[1:] if any(r)]
        session["activity_import_data"] = {
            "headers": headers,
            "rows": [[("" if v is None else str(v).strip()) for v in r] for r in data_rows],
        }
        return render_template("admin/activity_import.html",
                               preview={
                                   "headers": headers,
                                   "rows": [list(r) for r in data_rows[:20]],
                                   "total": len(data_rows),
                                   "mapping": mapping,
                               },
                               error=None,
                               projects=Project.query.filter_by(active=True).order_by(Project.sort_order).all())

    def _download_activity_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "活动参与名单"
        headers = ["项目", "日期", "时间", "地点", "活动内容", "服务对象人数",
                   "姓名", "学号", "A类小时数", "B类小时数", "学期"]
        header_fill = PatternFill("solid", start_color="4472C4")
        header_font = Font(bold=True, color="FFFFFF")
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = header_fill
            c.font = header_font
        sample = [
            ["心暖夕阳", "2025-11-15", "14:00-17:00", "夕阳红敬老院", "陪老人聊天", 20, "张三", "22301010001", 3.0, 0, "25秋冬"],
            ["心暖夕阳", "2025-11-15", "14:00-17:00", "夕阳红敬老院", "陪老人聊天", 20, "李四", "22301010002", 2.5, 0.5, "25秋冬"],
            ["心暖夕阳", "2025-11-15", "14:00-17:00", "夕阳红敬老院", "陪老人聊天", 20, "王五", "22301010003", 0, 2.0, "25秋冬"],
        ]
        for r, row in enumerate(sample, 2):
            for i, v in enumerate(row, 1):
                ws.cell(row=r, column=i, value=v)
        # 说明
        ws.cell(row=7, column=1, value="说明：")
        ws.cell(row=8, column=1, value="1. A类小时数 = 公益服务类（积分 ×1.0），B类小时数 = 事务类（积分 ×0.5）")
        ws.cell(row=9, column=1, value="2. 必填项：项目、姓名、（A或B类小时数至少一项）；同一日期+地点+内容的记录自动合并为同一场活动")
        ws.cell(row=10, column=1, value="3. 社员匹配优先级：学号 > 姓名")
        for i, w in enumerate([12, 12, 12, 16, 16, 8, 10, 14, 12, 12, 8], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="爱心社活动批量导入模板.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.route("/admin/activities/<int:aid>/remove_participant/<int:pid>", methods=["POST"])
    @permission_required("activities")
    def admin_activity_remove_participant(aid, pid):
        ap = ActivityParticipant.query.filter_by(activity_id=aid, member_id=pid).first_or_404()
        db.session.delete(ap)
        db.session.commit()
        flash("已移除参与者", "warning")
        return redirect(url_for("admin_activity_detail", aid=aid))

    # ---------- 任职积分 ----------
    @app.route("/admin/positions")
    @permission_required("positions")
    def admin_positions():
        sem = request.args.get("semester", current_semester())
        records = PositionRecord.query.filter_by(semester=sem).order_by(PositionRecord.id.desc()).all()
        return render_template("admin/positions.html", records=records, current_sem=sem)

    @app.route("/admin/positions/new", methods=["GET", "POST"])
    @permission_required("positions")
    def admin_position_new():
        if request.method == "POST":
            mid = request.form.get("member_id", type=int)
            if not mid:
                flash("请选择社员", "danger")
                return redirect(url_for("admin_position_new"))
            p = PositionRecord(
                member_id=mid,
                position=(request.form.get("position") or "").strip(),
                department=(request.form.get("department") or "").strip(),
                points=float(request.form.get("points") or 0),
                semester=(request.form.get("semester") or current_semester()).strip(),
            )
            db.session.add(p)
            db.session.commit()
            flash("已添加任职积分", "success")
            return redirect(url_for("admin_positions"))
        members = Member.query.order_by(Member.name).all()
        return render_template("admin/position_edit.html", record=None,
                               members=members, departments=all_departments())

    @app.route("/admin/positions/<int:pid>/delete", methods=["POST"])
    @permission_required("positions")
    def admin_position_delete(pid):
        p = PositionRecord.query.get_or_404(pid)
        db.session.delete(p)
        db.session.commit()
        flash("已删除", "warning")
        return redirect(url_for("admin_positions"))

    # ---------- 任职积分批量导入 ----------
    POSITION_COL_ALIASES = {
        "name": ["姓名", "name", "社员", "社员姓名"],
        "student_id": ["学号", "sid", "student_id"],
        "position": ["职务", "position", "岗位", "职位"],
        "department": ["部门", "department", "dept", "所属部门"],
        "points": ["积分", "points", "分值", "分数"],
        "semester": ["学期", "semester"],
    }

    @app.route("/admin/positions/import", methods=["GET", "POST"])
    @permission_required("positions")
    def admin_position_import():
        if request.method == "GET":
            return render_template("admin/position_import.html",
                                   preview=None, error=None,
                                   departments=all_departments())

        action = request.form.get("action", "preview")

        if action == "download_template":
            return _download_position_template()

        if action == "confirm":
            cached = session.get("position_import_data")
            if not cached:
                return render_template("admin/position_import.html",
                                       preview=None, error="会话已过期，请重新上传",
                                       departments=all_departments())
            rows = cached["rows"]
            try:
                col_name = int(request.form.get("col_name", -1))
                col_sid = int(request.form.get("col_sid", -1))
                col_position = int(request.form.get("col_position", -1))
                col_dept = int(request.form.get("col_dept", -1))
                col_points = int(request.form.get("col_points", -1))
                col_semester = int(request.form.get("col_semester", -1))
                default_semester = request.form.get("default_semester", current_semester()).strip()
            except (ValueError, TypeError):
                return render_template("admin/position_import.html",
                                       preview=None, error="列对应配置错误",
                                       departments=all_departments())

            if col_name < 0 and col_sid < 0:
                return render_template("admin/position_import.html",
                                       preview=None, error="必须指定「姓名」或「学号」列",
                                       departments=all_departments())
            if col_position < 0:
                return render_template("admin/position_import.html",
                                       preview=None, error="必须指定「职务」列",
                                       departments=all_departments())

            def get_col(r, idx):
                if idx < 0 or idx >= len(r):
                    return ""
                v = r[idx]
                return str(v).strip() if v is not None else ""

            members_by_sid = {m.student_id: m for m in Member.query.all() if m.student_id}
            members_by_name = {m.name: m for m in Member.query.all()}

            success, skipped, failed = 0, 0, 0
            failed_details = []

            for idx, r in enumerate(rows, 1):
                try:
                    sid = get_col(r, col_sid)
                    name = get_col(r, col_name)
                    member = None
                    if sid and sid in members_by_sid:
                        member = members_by_sid[sid]
                    elif name and name in members_by_name:
                        member = members_by_name[name]
                    if not member:
                        failed += 1
                        failed_details.append(f"第{idx}行：未找到社员【{name or sid}】")
                        continue

                    position = get_col(r, col_position)
                    if not position:
                        failed += 1
                        failed_details.append(f"第{idx}行：职务为空")
                        continue

                    points_v = get_col(r, col_points)
                    try:
                        points = float(points_v) if points_v else 0.0
                    except ValueError:
                        points = 0.0

                    semester = get_col(r, col_semester) or default_semester
                    dept = get_col(r, col_dept) or member.department or ""

                    # 重复检查：同一人同一学期同一职务视为重复
                    existing = PositionRecord.query.filter_by(
                        member_id=member.id, position=position, semester=semester).first()
                    if existing:
                        skipped += 1
                        continue

                    p = PositionRecord(
                        member_id=member.id,
                        position=position,
                        department=dept,
                        points=points,
                        semester=semester,
                    )
                    db.session.add(p)
                    success += 1
                except Exception as e:
                    failed += 1
                    failed_details.append(f"第{idx}行：{e}")

            if success > 0:
                log_modification(session.get("admin_name", "管理员"), None,
                                 f"批量导入社任职积分：成功 {success} 条，跳过 {skipped} 条，失败 {failed} 条")
            db.session.commit()
            session.pop("position_import_data", None)

            flash(f"导入完成：✅ 成功 {success} 条，跳过重复 {skipped} 条，失败 {failed} 条", "success")
            if failed_details:
                flash("失败明细：" + "；".join(failed_details[:10]), "warning")
            return redirect(url_for("admin_positions"))

        # preview
        f = request.files.get("file")
        if not f or not f.filename:
            return render_template("admin/position_import.html",
                                   preview=None, error="请选择文件",
                                   departments=all_departments())
        filename = f.filename.lower()
        try:
            if filename.endswith(".csv"):
                content = f.read().decode("utf-8-sig")
                reader = csv.reader(io.StringIO(content))
                rows = list(reader)
            elif filename.endswith((".xlsx", ".xls")):
                wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True, read_only=True)
                ws = wb.active
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                wb.close()
            else:
                return render_template("admin/position_import.html",
                                       preview=None, error="仅支持 .xlsx/.xls/.csv",
                                       departments=all_departments())
        except Exception as e:
            return render_template("admin/position_import.html",
                                   preview=None, error=f"文件解析失败：{e}",
                                   departments=all_departments())

        if not rows:
            return render_template("admin/position_import.html",
                                   preview=None, error="文件为空",
                                   departments=all_departments())

        headers = [str(c or "").strip() for c in rows[0]]
        mapping = {}
        used = set()
        for field, aliases in POSITION_COL_ALIASES.items():
            for i, h in enumerate(headers):
                if i in used:
                    continue
                h_clean = str(h or "").strip().lower().replace(" ", "")
                if h_clean in [a.lower().replace(" ", "") for a in aliases]:
                    mapping[field] = i
                    used.add(i)
                    break
        data_rows = [r for r in rows[1:] if any(r)]
        session["position_import_data"] = {
            "headers": headers,
            "rows": [[("" if v is None else str(v).strip()) for v in r] for r in data_rows],
        }
        return render_template("admin/position_import.html",
                               preview={
                                   "headers": headers,
                                   "rows": [list(r) for r in data_rows[:20]],
                                   "total": len(data_rows),
                                   "mapping": mapping,
                               },
                               error=None, departments=all_departments())

    def _download_position_template():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "任职积分名单"
        headers = ["姓名", "学号", "职务", "部门", "积分", "学期"]
        header_fill = PatternFill("solid", start_color="4472C4")
        header_font = Font(bold=True, color="FFFFFF")
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = header_fill
            c.font = header_font
        sample = [
            ["张三", "22301010001", "社长", "", 25, "25秋冬"],
            ["李四", "22301010002", "对内副社长", "", 18, "25秋冬"],
            ["王五", "22301010003", "人资部部长", "人资部", 12, "25秋冬"],
            ["赵六", "22301010004", "活动部副部长", "活动部", 8, "25秋冬"],
        ]
        for r, row in enumerate(sample, 2):
            for i, v in enumerate(row, 1):
                ws.cell(row=r, column=i, value=v)
        # 说明
        ws.cell(row=8, column=1, value="说明：")
        ws.cell(row=9, column=1, value="1. 必填项：姓名/学号（二选一）、职务")
        ws.cell(row=10, column=1, value="2. 部门可留空（默认与社员所属部门一致）")
        ws.cell(row=11, column=1, value="3. 学期为空时使用默认学期")
        ws.cell(row=12, column=1, value="4. 社员匹配优先级：学号 > 姓名")
        ws.cell(row=13, column=1, value="5. 同一人+同学期+同职务视为重复，将自动跳过")
        for i, w in enumerate([12, 14, 16, 12, 8, 8], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="爱心社任职积分导入模板.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


    # ---------- 非常规活动（事件 + 参与者） ----------
    @app.route("/admin/irregulars")
    @permission_required("irregulars")
    def admin_irregulars():
        records = IrregularActivity.query.order_by(IrregularActivity.id.desc()).all()
        return render_template("admin/irregulars.html", records=records)

    @app.route("/admin/irregulars/new", methods=["GET", "POST"])
    @permission_required("irregulars")
    def admin_irregular_new():
        if request.method == "POST":
            ia = IrregularActivity(
                content=(request.form.get("content") or "").strip(),
                date=(request.form.get("date") or "").strip(),
                location=(request.form.get("location") or "").strip(),
                a_hours=float(request.form.get("a_hours") or 0),
                b_hours=float(request.form.get("b_hours") or 0),
                semester=(request.form.get("semester") or current_semester()).strip(),
            )
            db.session.add(ia)
            db.session.commit()
            # 批量参与者
            try:
                pdata = json.loads(request.form.get("participants_data") or "[]")
            except Exception:
                pdata = []
            added = 0
            for item in pdata:
                mid = item.get("member_id")
                if not mid:
                    continue
                a_h = float(item.get("a_hours") or 0)
                b_h = float(item.get("b_hours") or 0)
                if a_h <= 0 and b_h <= 0:
                    continue
                if Member.query.get(mid) is None:
                    continue
                db.session.add(IrregularParticipant(
                    irregular_id=ia.id, member_id=mid,
                    a_hours=a_h, b_hours=b_h,
                ))
                added += 1
            db.session.commit()
            flash(f"已添加非常规活动：{ia.content or ia.date}（{added} 位参与者）", "success")
            return redirect(url_for("admin_irregular_detail", iid=ia.id))
        members = Member.query.order_by(Member.name).all()
        return render_template("admin/irregular_edit.html", activity=None,
                               members=members, participants=[])

    @app.route("/admin/irregulars/<int:iid>")
    @permission_required("irregulars")
    def admin_irregular_detail(iid):
        ia = IrregularActivity.query.get_or_404(iid)
        members = Member.query.order_by(Member.name).all()
        return render_template("admin/irregular_detail.html", activity=ia, members=members)

    @app.route("/admin/irregulars/<int:iid>/edit", methods=["GET", "POST"])
    @permission_required("irregulars")
    def admin_irregular_edit(iid):
        ia = IrregularActivity.query.get_or_404(iid)
        if request.method == "POST":
            for f in ["content", "date", "location", "semester"]:
                ia.__setattr__(f, (request.form.get(f) or "").strip())
            ia.a_hours = float(request.form.get("a_hours") or 0)
            ia.b_hours = float(request.form.get("b_hours") or 0)
            # 同步参与者
            try:
                pdata = json.loads(request.form.get("participants_data") or "[]")
            except Exception:
                pdata = []
            existing = {p.member_id: p for p in ia.participants}
            sent_ids = set()
            for item in pdata:
                mid = item.get("member_id")
                if not mid:
                    continue
                a_h = float(item.get("a_hours") or 0)
                b_h = float(item.get("b_hours") or 0)
                if a_h <= 0 and b_h <= 0:
                    continue
                if Member.query.get(mid) is None:
                    continue
                sent_ids.add(int(mid))
                if int(mid) in existing:
                    existing[int(mid)].a_hours = a_h
                    existing[int(mid)].b_hours = b_h
                else:
                    db.session.add(IrregularParticipant(
                        irregular_id=ia.id, member_id=int(mid),
                        a_hours=a_h, b_hours=b_h,
                    ))
            for mid, p in existing.items():
                if mid not in sent_ids:
                    db.session.delete(p)
            db.session.commit()
            flash("非常规活动已更新", "success")
            return redirect(url_for("admin_irregular_detail", iid=ia.id))
        members = Member.query.order_by(Member.name).all()
        participants = ia.participants
        return render_template("admin/irregular_edit.html", activity=ia,
                               members=members, participants=participants)

    @app.route("/admin/irregulars/<int:iid>/add_participant", methods=["POST"])
    @permission_required("irregulars")
    def admin_irregular_add_participant(iid):
        ia = IrregularActivity.query.get_or_404(iid)
        mid = request.form.get("member_id", type=int)
        a_h = float(request.form.get("a_hours") or 0)
        b_h = float(request.form.get("b_hours") or 0)
        if not mid or (a_h <= 0 and b_h <= 0):
            flash("请选择社员并填写 A 类或 B 类小时数", "danger")
            return redirect(url_for("admin_irregular_detail", iid=iid))
        existing = IrregularParticipant.query.filter_by(irregular_id=iid, member_id=mid).first()
        if existing:
            existing.a_hours = a_h
            existing.b_hours = b_h
        else:
            db.session.add(IrregularParticipant(
                irregular_id=iid, member_id=mid,
                a_hours=a_h, b_hours=b_h,
            ))
        db.session.commit()
        flash("已添加参与者", "success")
        return redirect(url_for("admin_irregular_detail", iid=iid))

    @app.route("/admin/irregulars/<int:iid>/remove_participant/<int:pid>", methods=["POST"])
    @permission_required("irregulars")
    def admin_irregular_remove_participant(iid, pid):
        p = IrregularParticipant.query.filter_by(irregular_id=iid, member_id=pid).first_or_404()
        db.session.delete(p)
        db.session.commit()
        flash("已移除参与者", "warning")
        return redirect(url_for("admin_irregular_detail", iid=iid))

    @app.route("/admin/irregulars/<int:iid>/delete", methods=["POST"])
    @permission_required("irregulars")
    def admin_irregular_delete(iid):
        ia = IrregularActivity.query.get_or_404(iid)
        db.session.delete(ia)
        db.session.commit()
        flash("已删除非常规活动", "warning")
        return redirect(url_for("admin_irregulars"))

    # ---------- 长期项目组 ----------
    @app.route("/admin/long_term_groups")
    @permission_required("long_term_groups")
    def admin_long_term_groups():
        groups = LongTermGroup.query.order_by(LongTermGroup.sort_order).all()
        # 统计每个组的小时数（来自所有关联成员的活动 + 非常规 + 任职）
        group_stats = []
        for g in groups:
            a_h = b_h = 0.0
            pos_pts = 0.0
            for m in g.members:
                a_h += m.a_hours()
                b_h += m.b_hours()
                for p in m.position_records:
                    pos_pts += p.points or 0
            group_stats.append({
                "group": g,
                "member_count": len(g.members),
                "a_hours": round(a_h, 1),
                "b_hours": round(b_h, 1),
                "total_points": round(a_h + b_h / 2.0, 2),
                "pos_points": round(pos_pts, 1),
            })
        return render_template("admin/long_term_groups.html", group_stats=group_stats)

    @app.route("/admin/long_term_groups/new", methods=["POST"])
    @permission_required("long_term_groups")
    def admin_long_term_group_new():
        name = (request.form.get("name") or "").strip()
        if not name:
            flash("组名不能为空", "danger")
            return redirect(url_for("admin_long_term_groups"))
        g = LongTermGroup(
            name=name,
            description=(request.form.get("description") or "").strip(),
            color=(request.form.get("color") or "#ff6b9d").strip(),
            sort_order=int(request.form.get("sort_order") or 0),
        )
        db.session.add(g)
        db.session.commit()
        log_modification(session.get("admin_name", "管理员"), None, f"新增长期项目组：{g.name}")
        db.session.commit()
        flash(f"已添加长期项目组【{g.name}】", "success")
        return redirect(url_for("admin_long_term_groups"))

    @app.route("/admin/long_term_groups/<int:gid>/edit", methods=["POST"])
    @permission_required("long_term_groups")
    def admin_long_term_group_edit(gid):
        g = LongTermGroup.query.get_or_404(gid)
        changes = []
        for field in ["name", "description", "color"]:
            new = (request.form.get(field) or "").strip()
            old = getattr(g, field) or ""
            if new != old:
                changes.append(f"字段【{field}】变更：\n  - 原值：{old}\n  - 新值：{new}")
                setattr(g, field, new)
        new_order = int(request.form.get("sort_order") or 0)
        if new_order != g.sort_order:
            changes.append(f"字段【sort_order】变更：\n  - 原值：{g.sort_order}\n  - 新值：{new_order}")
            g.sort_order = new_order
        new_active = request.form.get("active") == "1"
        if new_active != g.active:
            changes.append(f"字段【active】变更：\n  - 原值：{g.active}\n  - 新值：{new_active}")
            g.active = new_active
        if changes:
            log_modification(session.get("admin_name", "管理员"), None, "长期项目组【%s】修改：\n\n%s" % (g.name, "\n\n".join(changes)))
        db.session.commit()
        flash("保存成功", "success")
        return redirect(url_for("admin_long_term_groups"))

    @app.route("/admin/long_term_groups/<int:gid>/delete", methods=["POST"])
    @permission_required("long_term_groups")
    def admin_long_term_group_delete(gid):
        g = LongTermGroup.query.get_or_404(gid)
        # 将关联成员的 long_term_group_id 设为 None
        Member.query.filter_by(long_term_group_id=gid).update({"long_term_group_id": None})
        name = g.name
        db.session.delete(g)
        db.session.commit()
        log_modification(session.get("admin_name", "管理员"), None, f"删除长期项目组：{name}")
        db.session.commit()
        flash(f"已删除长期项目组【{name}】", "info")
        return redirect(url_for("admin_long_term_groups"))

    # ---------- 公开：长期项目组小时数公布 ----------
    @app.route("/long_term_groups")
    def long_term_groups_public():
        groups = LongTermGroup.query.filter_by(active=True).order_by(LongTermGroup.sort_order).all()
        group_stats = []
        for g in groups:
            a_h = b_h = 0.0
            pos_pts = 0.0
            for m in g.members:
                a_h += m.a_hours()
                b_h += m.b_hours()
                for p in m.position_records:
                    pos_pts += p.points or 0
            # 组员按总积分降序
            members_sorted = sorted(
                g.members, key=lambda x: x.total_points(), reverse=True
            )
            group_stats.append({
                "group": g,
                "member_count": len(g.members),
                "a_hours": round(a_h, 1),
                "b_hours": round(b_h, 1),
                "total_points": round(a_h + b_h / 2.0, 2),
                "pos_points": round(pos_pts, 1),
                "members": members_sorted,
            })
        # 按总积分降序排组
        group_stats.sort(key=lambda x: x["total_points"], reverse=True)
        return render_template("long_term_groups_public.html",
                               group_stats=group_stats,
                               now=datetime.now().strftime("%Y-%m-%d %H:%M"))

    # ---------- 黑名单 ----------
    @app.route("/admin/blacklist")
    @permission_required("blacklist")
    def admin_blacklist():
        records = Blacklist.query.order_by(Blacklist.id.desc()).all()
        return render_template("admin/blacklist.html", records=records)

    @app.route("/admin/blacklist/new", methods=["POST"])
    @permission_required("blacklist")
    def admin_blacklist_new():
        b = Blacklist(
            name=(request.form.get("name") or "").strip(),
            student_id=(request.form.get("student_id") or "").strip(),
            reason=(request.form.get("reason") or "").strip(),
        )
        db.session.add(b)
        db.session.commit()
        flash("已加入黑名单", "warning")
        return redirect(url_for("admin_blacklist"))

    @app.route("/admin/blacklist/<int:bid>/delete", methods=["POST"])
    @permission_required("blacklist")
    def admin_blacklist_delete(bid):
        b = Blacklist.query.get_or_404(bid)
        db.session.delete(b)
        db.session.commit()
        flash("已移出黑名单", "success")
        return redirect(url_for("admin_blacklist"))

    # ---------- 统计页 ----------
    @app.route("/admin/stats")
    @permission_required("stats")
    def admin_stats():
        sem = request.args.get("semester", "")
        members = Member.query.all()
        rows = []
        for m in members:
            rows.append({
                "member": m,
                "total": m.total_points(),
                "a_hours": m.a_hours(),
                "b_hours": m.b_hours(),
                "total_hours": m.total_hours(),
            })
        rows.sort(key=lambda x: -x["total"])
        # 部门积分（仅来自任职）
        dept_points = {}
        for m in members:
            for pr in m.position_records:
                d = pr.department or "未分配"
                dept_points[d] = dept_points.get(d, 0) + (pr.points or 0)
        # 项目 A/B 时长
        proj_a = {}
        proj_b = {}
        for p in Project.query.all():
            a = db.session.query(func.sum(ActivityParticipant.a_hours))\
                .join(Activity).filter(Activity.project_id == p.id).scalar() or 0
            b = db.session.query(func.sum(ActivityParticipant.b_hours))\
                .join(Activity).filter(Activity.project_id == p.id).scalar() or 0
            proj_a[p.name] = round(float(a), 1)
            proj_b[p.name] = round(float(b), 1)
        # 全局 A/B 时长
        total_a = sum(r["a_hours"] for r in rows)
        total_b = sum(r["b_hours"] for r in rows)
        return render_template("admin/stats.html",
                               rows=rows, dept_points=dept_points,
                               proj_a=proj_a, proj_b=proj_b,
                               total_a=round(total_a, 1), total_b=round(total_b, 1),
                               current_sem=sem)

    # ---------- 审计日志 ----------
    @app.route("/admin/logs")
    @permission_required("logs")
    def admin_logs():
        queries = QueryLog.query.order_by(QueryLog.id.desc()).all()
        signoffs = SignoffLog.query.order_by(SignoffLog.id.desc()).all()
        mods = ModificationLog.query.order_by(ModificationLog.id.desc()).all()
        me = Admin.query.get(session.get("admin_id"))
        return render_template("admin/logs.html", queries=queries,
                               signoffs=signoffs, mods=mods, me=me)

    # ---------- 审计日志删除 ----------
    @app.route("/admin/logs/query/<int:lid>/delete", methods=["POST"])
    @admin_required
    def admin_log_query_delete(lid):
        q = QueryLog.query.get_or_404(lid)
        info = f"id={q.id} {q.queryer_name or '?'}（学号 {q.query_code[:4]}***）"
        db.session.delete(q)
        db.session.commit()
        log_modification(session["admin_name"], None, f"删除查询记录：{info}")
        db.session.commit()
        flash(f"已删除查询记录 #{q.id if q else lid}", "warning")
        return redirect(url_for("admin_logs"))

    @app.route("/admin/logs/signoff/<int:lid>/delete", methods=["POST"])
    @admin_required
    def admin_log_signoff_delete(lid):
        s = SignoffLog.query.get_or_404(lid)
        info = f"id={s.id} {s.signoff_member.name if s.signoff_member else '?'}"
        db.session.delete(s)
        db.session.commit()
        log_modification(session["admin_name"], None, f"删除签收记录：{info}")
        db.session.commit()
        flash(f"已删除签收记录 #{s.id if s else lid}", "warning")
        return redirect(url_for("admin_logs"))

    @app.route("/admin/logs/modification/<int:lid>/delete", methods=["POST"])
    @admin_required
    def admin_log_modification_delete(lid):
        m = ModificationLog.query.get_or_404(lid)
        info = f"id={m.id} {m.modifier or '?'}"
        db.session.delete(m)
        db.session.commit()
        log_modification(session["admin_name"], None, f"删除修改记录：{info}")
        db.session.commit()
        flash(f"已删除修改记录 #{m.id if m else lid}", "warning")
        return redirect(url_for("admin_logs"))

    # ---------- 审计日志批量删除 ----------
    @app.route("/admin/logs/query/batch_delete", methods=["POST"])
    @admin_required
    def admin_log_query_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请至少选择一条记录", "warning")
            return redirect(url_for("admin_logs"))
        try:
            ids = [int(i) for i in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_logs"))
        deleted = QueryLog.query.filter(QueryLog.id.in_(ids)).delete(synchronize_session=False)
        db.session.commit()
        log_modification(session["admin_name"], None, f"批量删除查询记录：{deleted} 条")
        db.session.commit()
        flash(f"已批量删除 {deleted} 条查询记录", "warning")
        return redirect(url_for("admin_logs"))

    @app.route("/admin/logs/signoff/batch_delete", methods=["POST"])
    @admin_required
    def admin_log_signoff_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请至少选择一条记录", "warning")
            return redirect(url_for("admin_logs"))
        try:
            ids = [int(i) for i in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_logs"))
        deleted = SignoffLog.query.filter(SignoffLog.id.in_(ids)).delete(synchronize_session=False)
        db.session.commit()
        log_modification(session["admin_name"], None, f"批量删除签收记录：{deleted} 条")
        db.session.commit()
        flash(f"已批量删除 {deleted} 条签收记录", "warning")
        return redirect(url_for("admin_logs"))

    @app.route("/admin/logs/modification/batch_delete", methods=["POST"])
    @admin_required
    def admin_log_modification_batch_delete():
        ids = request.form.getlist("ids")
        if not ids:
            flash("请至少选择一条记录", "warning")
            return redirect(url_for("admin_logs"))
        try:
            ids = [int(i) for i in ids]
        except ValueError:
            flash("参数错误", "danger")
            return redirect(url_for("admin_logs"))
        deleted = ModificationLog.query.filter(ModificationLog.id.in_(ids)).delete(synchronize_session=False)
        db.session.commit()
        log_modification(session["admin_name"], None, f"批量删除修改记录：{deleted} 条")
        db.session.commit()
        flash(f"已批量删除 {deleted} 条修改记录", "warning")
        return redirect(url_for("admin_logs"))

    # ---------- Excel 导出 ----------
    @app.route("/admin/export")
    @permission_required("export")
    def admin_export():
        wb = openpyxl.Workbook()

        # 工具函数
        header_fill = PatternFill("solid", start_color="4472C4")
        header_font = Font(bold=True, color="FFFFFF", name="Arial")
        thin = Side(border_style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        def write_header(ws, headers):
            for i, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=i, value=h)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center")
                c.border = border

        def auto_width(ws):
            for col in ws.columns:
                m = 8
                letter = get_column_letter(col[0].column)
                for c in col:
                    v = c.value
                    if v is not None:
                        m = max(m, len(str(v)) + 2)
                ws.column_dimensions[letter].width = min(m, 30)

        # Sheet 1: 积分总表（按 A/B 分类）
        ws = wb.active
        ws.title = "积分总表"
        projects = Project.query.filter_by(active=True).order_by(Project.sort_order).all()
        headers = ["姓名", "年级", "学号", "部门", "A类小时数", "B类小时数",
                   "总小时数", "总积分(A+B/2)", "任职积分"]
        write_header(ws, headers)
        for r, m in enumerate(Member.query.order_by(Member.id).all(), 2):
            ws.cell(row=r, column=1, value=m.name)
            ws.cell(row=r, column=2, value=m.grade)
            ws.cell(row=r, column=3, value=m.student_id)
            ws.cell(row=r, column=4, value=m.department)
            ws.cell(row=r, column=5, value=m.a_hours())
            ws.cell(row=r, column=6, value=m.b_hours())
            ws.cell(row=r, column=7, value=m.total_hours())
            ws.cell(row=r, column=8, value=f"=E{r}+F{r}/2")
            pos_pts = sum(p.points or 0 for p in m.position_records)
            ws.cell(row=r, column=9, value=pos_pts)
        auto_width(ws)

        # Sheet 2: 全部活动总表
        ws = wb.create_sheet("活动总表")
        write_header(ws, ["项目", "活动次序", "日期", "时间", "地点",
                          "活动内容", "服务对象人数", "志愿者", "学号", "部门",
                          "A类小时数", "B类小时数", "贡献积分", "学期"])
        r = 2
        for a in Activity.query.order_by(Activity.date, Activity.id).all():
            base = [a.project.name, a.sequence, a.date, a.time, a.location,
                    a.content, a.served_count]
            if a.participants:
                for ap in a.participants:
                    pts = (ap.a_hours or 0) + (ap.b_hours or 0) / 2.0
                    row = base + [ap.member.name, ap.member.student_id, ap.member.department,
                                  ap.a_hours, ap.b_hours, round(pts, 2), a.semester]
                    for i, v in enumerate(row, 1):
                        ws.cell(row=r, column=i, value=v)
                    r += 1
            else:
                row = base + ["", "", "", "", "", "", a.semester]
                for i, v in enumerate(row, 1):
                    ws.cell(row=r, column=i, value=v)
                r += 1
        auto_width(ws)

        # Sheet 3: 各项目分表
        for p in projects:
            ws = wb.create_sheet(p.name)
            write_header(ws, ["活动次序", "日期", "时间", "地点", "活动内容",
                              "服务对象人数", "志愿者", "学号", "部门", "手机号",
                              "A类小时数", "B类小时数", "贡献积分", "学期"])
            r = 2
            for a in p.activities:
                base = [a.sequence, a.date, a.time, a.location, a.content, a.served_count]
                if a.participants:
                    for ap in a.participants:
                        pts = (ap.a_hours or 0) + (ap.b_hours or 0) / 2.0
                        row = base + [ap.member.name, ap.member.student_id,
                                      ap.member.department, ap.member.phone,
                                      ap.a_hours, ap.b_hours, round(pts, 2), a.semester]
                        for i, v in enumerate(row, 1):
                            ws.cell(row=r, column=i, value=v)
                        r += 1
                else:
                    row = base + ["", "", "", "", "", "", "", a.semester]
                    for i, v in enumerate(row, 1):
                        ws.cell(row=r, column=i, value=v)
                    r += 1
            auto_width(ws)

        # 任职积分
        ws = wb.create_sheet("任职积分")
        write_header(ws, ["姓名", "职务", "部门", "积分", "学期"])
        for r, p in enumerate(PositionRecord.query.order_by(PositionRecord.id).all(), 2):
            ws.cell(row=r, column=1, value=p.member.name)
            ws.cell(row=r, column=2, value=p.position)
            ws.cell(row=r, column=3, value=p.department)
            ws.cell(row=r, column=4, value=p.points)
            ws.cell(row=r, column=5, value=p.semester)
        auto_width(ws)

        # 非常规活动
        ws = wb.create_sheet("非常规活动")
        write_header(ws, ["活动内容", "日期", "地点", "志愿者", "学号", "部门",
                          "A类小时数", "B类小时数", "贡献积分", "学期"])
        r = 2
        for ia in IrregularActivity.query.order_by(IrregularActivity.id).all():
            base = [ia.content, ia.date, ia.location]
            if ia.participants:
                for ip in ia.participants:
                    pts = (ip.a_hours or 0) + (ip.b_hours or 0) / 2.0
                    row = base + [ip.member.name, ip.member.student_id, ip.member.department,
                                  ip.a_hours, ip.b_hours, round(pts, 2), ia.semester]
                    for i, v in enumerate(row, 1):
                        ws.cell(row=r, column=i, value=v)
                    r += 1
            else:
                row = base + ["", "", "", "", "", "", ia.semester]
                for i, v in enumerate(row, 1):
                    ws.cell(row=r, column=i, value=v)
                r += 1
        auto_width(ws)

        # Sheet 13: 查询记录
        ws = wb.create_sheet("查询记录")
        write_header(ws, ["查询时间", "查询人", "学号", "关联社员"])
        for r, q in enumerate(QueryLog.query.order_by(QueryLog.id.desc()).all(), 2):
            ws.cell(row=r, column=1, value=q.query_time)
            ws.cell(row=r, column=2, value=q.queryer_name)
            ws.cell(row=r, column=3, value=q.query_code)
            ws.cell(row=r, column=4, value=q.query_member.name if q.query_member else "")
        auto_width(ws)

        # Sheet 14: 签收记录
        ws = wb.create_sheet("签收记录")
        write_header(ws, ["签收时间", "签收方式", "签收人", "学号"])
        for r, s in enumerate(SignoffLog.query.order_by(SignoffLog.id.desc()).all(), 2):
            ws.cell(row=r, column=1, value=s.signoff_time)
            ws.cell(row=r, column=2, value=s.method)
            ws.cell(row=r, column=3, value=s.signoff_member.name if s.signoff_member else "")
            ws.cell(row=r, column=4, value=s.related_query_code)
        auto_width(ws)

        # Sheet 15: 修改记录
        ws = wb.create_sheet("修改记录")
        write_header(ws, ["修改时间", "修改人", "关联社员", "修改明细", "学号"])
        member_map = {m.id: m for m in Member.query.all()}
        for r, m in enumerate(ModificationLog.query.order_by(ModificationLog.id.desc()).all(), 2):
            ws.cell(row=r, column=1, value=m.submit_time)
            ws.cell(row=r, column=2, value=m.modifier)
            mm = member_map.get(m.member_id) if m.member_id else None
            ws.cell(row=r, column=3, value=mm.name if mm else "")
            ws.cell(row=r, column=4, value=m.details)
            ws.cell(row=r, column=5, value=m.query_code)
        auto_width(ws)

        # 输出
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        fname = f"爱心社积分系统_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---------- 系统设置 ----------
    @app.route("/admin/settings", methods=["GET", "POST"])
    @permission_required("settings")
    def admin_settings():
        if request.method == "POST":
            year_raw = (request.form.get("current_academic_year") or "").strip()
            season = (request.form.get("current_season") or "").strip()
            sem = (request.form.get("current_semester") or "").strip()
            depts = (request.form.get("departments") or "").strip()
            # 解析学年：兼容 "2025" 或 "2025-2026" 两种格式
            if year_raw and season and not sem:
                try:
                    first_year = int(year_raw.split("-")[0])
                    if season == "秋冬":
                        short = f"{first_year % 100:02d}{season}"
                    else:
                        short = f"{(first_year + 1) % 100:02d}{season}"
                    sem = short
                except Exception:
                    pass
            for k, v in [("current_semester", sem), ("departments", depts)]:
                if v:
                    s = Setting.query.filter_by(key=k).first()
                    if not s:
                        s = Setting(key=k, value=v)
                        db.session.add(s)
                    else:
                        s.value = v
            db.session.commit()
            flash("设置已保存", "success")
            return redirect(url_for("admin_settings"))
        cur = get_setting("current_semester", "25秋冬")
        cur_year, cur_season = "", ""
        if cur and len(cur) >= 4:
            short_year = int(cur[:2])
            season = cur[2:]
            if season == "秋冬":
                cur_year = f"{short_year + 2000:04d}"
            else:
                cur_year = f"{short_year + 1999:04d}"
            cur_season = season
        import datetime as _dt
        this_year = _dt.date.today().year
        # 自动检测当前学年+学期（浙大规范）
        auto_year, auto_season = detect_current_semester()
        return render_template("admin/settings.html",
                               current_sem=cur,
                               cur_year=cur_year,
                               cur_season=cur_season,
                               auto_year=auto_year,
                               auto_season=auto_season,
                               departments=",".join(all_departments()))

    @app.errorhandler(404)
    def not_found(e):
        return render_template("error.html", code=404, msg="页面不存在"), 404

    @app.errorhandler(500)
    def server_error(e):
        return render_template("error.html", code=500, msg="服务器内部错误"), 500


app = create_app()


if __name__ == "__main__":
    print("=" * 50)
    print("♥  ZJU Club of Heart Caring")
    print("♥  浙江大学学生爱心社 · 志愿服务积分系统")
    print("=" * 50)
    print("  服务社会 · 奉献爱心 · 推己及人 · 薪火相传")
    print("=" * 50)
    print("📍 访问地址（本机）: http://localhost:5050")
    print("🍎 桌面/启动台/聚焦搜索：「ZJU Club of Heart Caring」")
    print("   ↳ 已安装 .app 到 /Applications/ + 桌面别名（双击即打开）")
    print("🔑 管理员密码: aixin2026")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=5050)
