"""数据模型 - ZJU Club of Heart Caring · 浙江大学学生爱心社
设计：
- 活动 (Activity) = 一次事件，下挂多个参与者
- 每个参与者分别记录 A类小时数 与 B类小时数（A 类公益 ×1.0，B 类事务 ×0.5）
- 活动本身可有 a_hours / b_hours 字段作为活动总时长参考（不必等于参与者总和）
- 非常规活动 (IrregularActivity) = 一次事件，结构同 Activity，无项目归属
- 非常规参与者 (IrregularParticipant) = 同 ActivityParticipant
"""
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Member(db.Model):
    """成员（社员 / 长期项目组成员）"""
    __tablename__ = "members"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), nullable=False, index=True)
    grade = db.Column(db.String(32), default="")
    student_id = db.Column(db.String(32), default="", index=True)
    department = db.Column(db.String(64), default="")
    # 身份：社员 / 长期项目组成员
    identity = db.Column(db.String(32), default="社员", index=True)
    # 长期项目组（可空；可对应 LongTermGroup.id；也可填自由文本）
    long_term_group_id = db.Column(db.Integer, db.ForeignKey("long_term_groups.id"), nullable=True, index=True)
    long_term_group = db.relationship("LongTermGroup", backref="members")
    phone = db.Column(db.String(32), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    activity_participations = db.relationship("ActivityParticipant", backref="member", cascade="all, delete-orphan")
    irregular_participations = db.relationship("IrregularParticipant", backref="member", cascade="all, delete-orphan")
    position_records = db.relationship("PositionRecord", backref="member", cascade="all, delete-orphan")
    signoffs = db.relationship("SignoffLog", backref="signoff_member", cascade="all, delete-orphan")
    query_logs = db.relationship("QueryLog", backref="query_member", cascade="all, delete-orphan")
    modifications = db.relationship("ModificationLog", backref="mod_member", cascade="all, delete-orphan")

    # -------- A/B 分类统计 --------
    def a_hours(self):
        """A 类总小时数（活动 + 非常规）"""
        total = 0.0
        for ap in self.activity_participations:
            total += ap.a_hours or 0
        for ip in self.irregular_participations:
            total += ip.a_hours or 0
        return round(total, 2)

    def b_hours(self):
        """B 类总小时数（活动 + 非常规）"""
        total = 0.0
        for ap in self.activity_participations:
            total += ap.b_hours or 0
        for ip in self.irregular_participations:
            total += ip.b_hours or 0
        return round(total, 2)

    def total_hours(self):
        return round(self.a_hours() + self.b_hours(), 2)

    def total_points(self):
        """总积分 = A类小时数 + B类小时数 / 2"""
        a = self.a_hours()
        b = self.b_hours()
        return round(a + b / 2.0, 2)

    def project_points(self, project_id):
        """保留接口：返回该项目下 a_hours / b_hours"""
        a = 0.0
        b = 0.0
        for ap in self.activity_participations:
            if ap.activity and ap.activity.project_id == project_id:
                a += ap.a_hours or 0
                b += ap.b_hours or 0
        return round(a, 2), round(b, 2)

    # -------- 签收统计 --------
    def unsigned_activity_count(self):
        """未签收的活动参与记录数（成员需在查询页面确认）"""
        return sum(1 for ap in self.activity_participations if ap.signed_off_at is None) + \
               sum(1 for ip in self.irregular_participations if ip.signed_off_at is None)

    def signed_activity_count(self):
        """已签收的参与记录数"""
        return sum(1 for ap in self.activity_participations if ap.signed_off_at is not None) + \
               sum(1 for ip in self.irregular_participations if ip.signed_off_at is not None)

    def latest_signoff_str(self):
        """最近一次签收时间（字符串 YYYY-MM-DD HH:MM），无则 None"""
        times = [t for t in (
            [ap.signed_off_at for ap in self.activity_participations if ap.signed_off_at] +
            [ip.signed_off_at for ip in self.irregular_participations if ip.signed_off_at] +
            [s.signoff_time for s in self.signoffs if s.signoff_time]
        ) if t is not None]
        if not times:
            return None
        return max(times).strftime('%Y-%m-%d %H:%M')


class Project(db.Model):
    """品牌项目（分类容器）"""
    __tablename__ = "projects"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), unique=True, nullable=False)
    description = db.Column(db.String(256), default="")
    sort_order = db.Column(db.Integer, default=0)
    active = db.Column(db.Boolean, default=True)

    activities = db.relationship("Activity", backref="project", cascade="all, delete-orphan")

    def activity_count(self):
        return len(self.activities)


class Activity(db.Model):
    """项目活动事件"""
    __tablename__ = "activities"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=False, index=True)
    sequence = db.Column(db.Integer, default=0)
    date = db.Column(db.String(32), default="")
    time = db.Column(db.String(32), default="")
    location = db.Column(db.String(128), default="")
    content = db.Column(db.String(256), default="")
    served_count = db.Column(db.Integer, default=0)
    # 活动本身的 A/B 时长（参考；不必等于参与者总和）
    a_hours = db.Column(db.Float, default=0.0)
    b_hours = db.Column(db.Float, default=0.0)
    semester = db.Column(db.String(16), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.String(64), default="")

    participants = db.relationship("ActivityParticipant", backref="activity", cascade="all, delete-orphan")

    def total_a_hours(self):
        return round(sum(ap.a_hours for ap in self.participants), 2)

    def total_b_hours(self):
        return round(sum(ap.b_hours for ap in self.participants), 2)

    def participant_count(self):
        return len(self.participants)

    def semester_seq(self):
        """在 (项目, 学期) 内按 date 升序的次序（1-based）"""
        if not self.semester:
            return 0
        siblings = (Activity.query
                    .filter_by(project_id=self.project_id, semester=self.semester)
                    .order_by(Activity.date.asc(), Activity.id.asc())
                    .all())
        for i, s in enumerate(siblings, start=1):
            if s.id == self.id:
                return i
        return 0


class ActivityParticipant(db.Model):
    """活动参与者（每人分别记录 A/B 小时数 + 单独签收）"""
    __tablename__ = "activity_participants"
    id = db.Column(db.Integer, primary_key=True)
    activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"), nullable=False, index=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False, index=True)
    a_hours = db.Column(db.Float, default=0.0)
    b_hours = db.Column(db.Float, default=0.0)
    note = db.Column(db.String(128), default="")
    signed_off_at = db.Column(db.DateTime, nullable=True)  # 成员确认签收时间
    __table_args__ = (db.UniqueConstraint("activity_id", "member_id", name="uq_act_member"),)

    def points(self):
        return round((self.a_hours or 0) + (self.b_hours or 0) / 2.0, 2)

    @property
    def is_signed_off(self):
        return self.signed_off_at is not None


class PositionRecord(db.Model):
    """任职积分"""
    __tablename__ = "position_records"
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False, index=True)
    position = db.Column(db.String(64), nullable=False)
    department = db.Column(db.String(64), default="")
    points = db.Column(db.Float, default=0.0)
    semester = db.Column(db.String(16), nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class IrregularActivity(db.Model):
    """非常规活动（事件）"""
    __tablename__ = "irregular_activities"
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.String(256), default="")
    date = db.Column(db.String(32), default="")
    location = db.Column(db.String(128), default="")
    semester = db.Column(db.String(16), default="")
    # 事件本身参考 A/B 时长
    a_hours = db.Column(db.Float, default=0.0)
    b_hours = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    participants = db.relationship("IrregularParticipant", backref="irregular", cascade="all, delete-orphan")

    def total_a_hours(self):
        return round(sum(p.a_hours for p in self.participants), 2)

    def total_b_hours(self):
        return round(sum(p.b_hours for p in self.participants), 2)

    def participant_count(self):
        return len(self.participants)


class IrregularParticipant(db.Model):
    """非常规活动参与者（每人分别记录 A/B 小时数 + 单独签收）"""
    __tablename__ = "irregular_participants"
    id = db.Column(db.Integer, primary_key=True)
    irregular_id = db.Column(db.Integer, db.ForeignKey("irregular_activities.id"), nullable=False, index=True)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False, index=True)
    a_hours = db.Column(db.Float, default=0.0)
    b_hours = db.Column(db.Float, default=0.0)
    note = db.Column(db.String(128), default="")
    signed_off_at = db.Column(db.DateTime, nullable=True)  # 成员确认签收时间
    __table_args__ = (db.UniqueConstraint("irregular_id", "member_id", name="uq_irreg_member"),)

    def points(self):
        return round((self.a_hours or 0) + (self.b_hours or 0) / 2.0, 2)

    @property
    def is_signed_off(self):
        return self.signed_off_at is not None


class QueryLog(db.Model):
    """查询记录"""
    __tablename__ = "query_log"
    id = db.Column(db.Integer, primary_key=True)
    query_time = db.Column(db.DateTime, default=datetime.utcnow)
    member_id = db.Column(db.Integer, db.ForeignKey("members.id", use_alter=True, name="fk_querylog_member"), nullable=True)
    queryer_name = db.Column(db.String(64), default="")
    query_code = db.Column(db.String(32), default="")
    source = db.Column(db.String(32), default="网页查询")


class SignoffLog(db.Model):
    """签收记录"""
    __tablename__ = "signoff_log"
    id = db.Column(db.Integer, primary_key=True)
    method = db.Column(db.String(64), default="点击确认按钮签收")
    member_id = db.Column(db.Integer, db.ForeignKey("members.id", use_alter=True, name="fk_signofflog_member"), nullable=False, index=True)
    signoff_time = db.Column(db.DateTime, default=datetime.utcnow)
    related_query_code = db.Column(db.String(32), default="")


class ModificationLog(db.Model):
    """修改记录"""
    __tablename__ = "modification_log"
    id = db.Column(db.Integer, primary_key=True)
    submit_time = db.Column(db.DateTime, default=datetime.utcnow)
    modifier = db.Column(db.String(64), default="")
    member_id = db.Column(db.Integer, db.ForeignKey("members.id", use_alter=True, name="fk_modificationlog_member"), nullable=True)
    details = db.Column(db.Text, default="")
    query_code = db.Column(db.String(32), default="")


class Blacklist(db.Model):
    """志愿者黑名单"""
    __tablename__ = "blacklist"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), default="")
    student_id = db.Column(db.String(32), default="")
    reason = db.Column(db.String(256), default="")
    added_at = db.Column(db.DateTime, default=datetime.utcnow)


class LongTermGroup(db.Model):
    """长期项目组（如唐奖组、年检组等）"""
    __tablename__ = "long_term_groups"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), nullable=False, unique=True)
    description = db.Column(db.String(256), default="")
    color = db.Column(db.String(16), default="#ff6b9d")  # 标识色
    sort_order = db.Column(db.Integer, default=0)
    active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Setting(db.Model):
    """系统设置（管理员密码等）"""
    __tablename__ = "settings"
    key = db.Column(db.String(64), primary_key=True)
    value = db.Column(db.String(256), default="")


class Admin(db.Model):
    """管理员白名单（多管理员 + 超管 + 创始人 + 细粒度权限 + 审批流）
    注册流程：
      1. 超级管理员可在 /admin/whitelist 预添加（无密码 = 待激活）→ 成员自行完成注册（设置密码）
      2. 成员可在 /admin/login 提交「申请注册」→ 创建待审批记录 → 超管在白名单页审批
      3. 一旦记录在白名单中且 password_hash 不为空 → 即可登录
    状态：
      - 'pending'   = 待审批（用户提交了申请，未被超管批准）
      - 'approved'  = 已批准（无密码 = 待激活；超管预添加 或 已审批未设置密码）
      - 'active'    = 活跃（有密码 = 可登录）
    """
    __tablename__ = "admins"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), nullable=False, index=True)
    student_id = db.Column(db.String(32), nullable=False, unique=True, index=True)
    password_hash = db.Column(db.String(256), default="")  # 允许空 = 待激活
    is_super_admin = db.Column(db.Boolean, default=False, index=True)
    is_root = db.Column(db.Boolean, default=False, index=True)  # 创始人：最高权限
    permissions = db.Column(db.Text, default="")  # JSON list, e.g. '["members","projects"]'
    status = db.Column(db.String(16), default="active", index=True)  # pending/approved/active
    application_message = db.Column(db.Text, default="")  # 申请说明
    approved_at = db.Column(db.DateTime, nullable=True)  # 审批时间
    approved_by = db.Column(db.String(64), default="")  # 审批人
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login_at = db.Column(db.DateTime, nullable=True)

    def is_activated(self):
        """是否可登录：active 状态 + 密码已设置"""
        return self.status == "active" and bool(self.password_hash)

    def is_pending_approval(self):
        return self.status == "pending"

    def is_awaiting_activation(self):
        """已批准但未设置密码"""
        return self.status == "approved" and not self.password_hash

    def set_password(self, raw):
        from werkzeug.security import generate_password_hash
        self.password_hash = generate_password_hash(raw)

    def check_password(self, raw):
        from werkzeug.security import check_password_hash
        return check_password_hash(self.password_hash, raw)

    def permissions_list(self):
        """返回权限列表（超管返回全部权限）"""
        if self.is_super_admin:
            return ALL_ADMIN_PERMISSIONS  # 全权限
        if not self.permissions:
            return []
        # 支持 JSON 数组和逗号分隔字符串
        s = self.permissions.strip()
        if s.startswith("["):
            import json
            try:
                return json.loads(s)
            except (ValueError, TypeError):
                return []
        return [p.strip() for p in s.split(",") if p.strip()]

    def has_permission(self, perm):
        """是否有某项权限"""
        if self.is_super_admin:
            return True
        return perm in self.permissions_list()


# 所有可分配的管理权限（超管默认拥有全部）
# 注意：白名单管理（whitelist）只能超管使用，永远不分配给普管
ALL_ADMIN_PERMISSIONS = [
    "dashboard",      # 后台首页
    "members",        # 成员管理
    "projects",       # 项目管理
    "activities",     # 活动管理
    "positions",      # 任职积分
    "irregulars",     # 非常规活动
    "long_term_groups",  # 长期项目组
    "blacklist",      # 黑名单
    "stats",          # 数据统计
    "logs",           # 审计日志
    "export",         # 导出
    "settings",       # 系统设置
    "profile",        # 我的资料
]
