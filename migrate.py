"""数据迁移：旧 schema -> 新 schema

- Activity.activity_type 删除，新增 a_hours/b_hours
- ActivityParticipant.service_hours 拆分为 a_hours/b_hours（按原活动类型归入对应列）
- IrregularActivity: 旧模型 (member_id, content, date, activity_type, hours) → 拆为
    * IrregularActivity(事件): content, date, semester
    * IrregularParticipant: member_id, a_hours, b_hours
"""
import os
import sys
from datetime import datetime
from flask import Flask
from sqlalchemy import inspect, text


def parse_dt(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    try:
        return datetime.fromisoformat(str(v))
    except Exception:
        return None

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from models import (
    db, Member, Project, Activity, ActivityParticipant,
    IrregularActivity, IrregularParticipant, PositionRecord,
    QueryLog, SignoffLog, ModificationLog, Blacklist, Setting,
)


def create_app():
    app = Flask(__name__)
    db_path = os.path.join(HERE, "data", "aixin.db")
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SECRET_KEY"] = "zju-aixin-secret-2026"
    db.init_app(app)
    return app


def fetch_old_data(conn, table):
    """从 conn 取一张表的所有列数据"""
    try:
        rows = conn.execute(text(f"SELECT * FROM {table}")).fetchall()
        cols = conn.execute(text(f"PRAGMA table_info({table})")).fetchall()
        col_names = [c[1] for c in cols]
        return [dict(zip(col_names, row)) for row in rows]
    except Exception as e:
        print(f"  读取 {table} 失败: {e}")
        return []


def migrate():
    app = create_app()
    with app.app_context():
        insp = inspect(db.engine)
        existing_tables = insp.get_table_names()
        print("现有表:", existing_tables)

        # 1. 在 drop 之前，把所有需要的数据全部读出
        backup = {}
        with db.engine.connect() as conn:
            for t in ["members", "projects", "activities", "activity_participants",
                      "position_records", "irregular_activities",
                      "query_log", "signoff_log", "modification_log",
                      "blacklist", "settings"]:
                backup[t] = fetch_old_data(conn, t)
                print(f"  备份 {t}: {len(backup[t])}")

        # 2. 重建 schema
        print("\n重建 schema...")
        db.drop_all()
        db.create_all()

        # 3. 写回基础数据（无 FK 依赖的先写）
        for r in backup["settings"]:
            db.session.add(Setting(key=r["key"], value=r.get("value", "")))
        for r in backup["members"]:
            db.session.add(Member(
                id=r["id"], name=r["name"],
                grade=r.get("grade", "") or "",
                student_id=r.get("student_id", "") or "",
                department=r.get("department", "") or "",
                query_code=r.get("query_code", "") or "",
                phone=r.get("phone", "") or "",
                created_at=parse_dt(r.get("created_at")),
            ))
        for r in backup["projects"]:
            db.session.add(Project(
                id=r["id"], name=r["name"],
                description=r.get("description", "") or "",
                sort_order=r.get("sort_order", 0) or 0,
                active=bool(r.get("active", True)),
            ))
        db.session.commit()
        print(f"  写入: members={len(backup['members'])}, projects={len(backup['projects'])}")

        # 4. 写回 Activity + 拆 service_hours
        a_type_map = {}  # activity_id -> 'A' / 'B'
        for r in backup["activities"]:
            db.session.add(Activity(
                id=r["id"], project_id=r["project_id"],
                sequence=r.get("sequence", 0) or 0,
                date=r.get("date", "") or "",
                time=r.get("time", "") or "",
                location=r.get("location", "") or "",
                content=r.get("content", "") or "",
                served_count=r.get("served_count", 0) or 0,
                semester=r.get("semester", "") or "",
                created_at=parse_dt(r.get("created_at")),
                created_by=r.get("created_by", "") or "",
            ))
            a_type_map[r["id"]] = r.get("activity_type", "A")
        db.session.commit()
        print(f"  写入: activities={len(backup['activities'])}")

        for r in backup["activity_participants"]:
            atype = a_type_map.get(r["activity_id"], "A")
            hrs = r.get("service_hours", 0) or 0.0
            a_h = hrs if atype == "A" else 0.0
            b_h = hrs if atype == "B" else 0.0
            db.session.add(ActivityParticipant(
                id=r["id"], activity_id=r["activity_id"], member_id=r["member_id"],
                a_hours=a_h, b_hours=b_h,
                note=r.get("note", "") or "",
            ))
        db.session.commit()
        print(f"  写入: activity_participants={len(backup['activity_participants'])}")

        # 5. 写回 IrregularActivity(事件) + IrregularParticipant
        for r in backup["irregular_activities"]:
            db.session.add(IrregularActivity(
                id=r["id"],
                content=r.get("content", "") or "",
                date=r.get("date", "") or "",
                semester=r.get("semester", "") or "",
            ))
        db.session.commit()
        print(f"  写入: irregular_activities={len(backup['irregular_activities'])}")

        # 把旧 IrregularActivity 拆为参与者
        for r in backup["irregular_activities"]:
            atype = r.get("activity_type", "B")
            hrs = r.get("hours", 0) or 0.0
            mid = r.get("member_id")
            if not mid:
                continue
            a_h = hrs if atype == "A" else 0.0
            b_h = hrs if atype == "B" else 0.0
            db.session.add(IrregularParticipant(
                irregular_id=r["id"], member_id=mid,
                a_hours=a_h, b_hours=b_h,
            ))
        db.session.commit()
        irregular_participants = len([r for r in backup["irregular_activities"] if r.get("member_id")])
        print(f"  写入: irregular_participants={irregular_participants}")

        # 6. 其他表
        for r in backup["position_records"]:
            db.session.add(PositionRecord(
                id=r["id"], member_id=r["member_id"],
                position=r["position"],
                department=r.get("department", "") or "",
                points=r.get("points", 0) or 0.0,
                semester=r["semester"],
                created_at=parse_dt(r.get("created_at")),
            ))
        db.session.commit()
        for r in backup["query_log"]:
            db.session.add(QueryLog(
                id=r["id"], query_time=parse_dt(r.get("query_time")),
                member_id=r.get("member_id"),
                queryer_name=r.get("queryer_name", "") or "",
                query_code=r.get("query_code", "") or "",
                source=r.get("source", "网页查询") or "网页查询",
            ))
        for r in backup["signoff_log"]:
            db.session.add(SignoffLog(
                id=r["id"], method=r.get("method", "点击确认按钮签收") or "",
                member_id=r["member_id"],
                signoff_time=parse_dt(r.get("signoff_time")),
                related_query_code=r.get("related_query_code", "") or "",
            ))
        for r in backup["modification_log"]:
            db.session.add(ModificationLog(
                id=r["id"], submit_time=parse_dt(r.get("submit_time")),
                modifier=r.get("modifier", "") or "",
                member_id=r.get("member_id"),
                details=r.get("details", "") or "",
                query_code=r.get("query_code", "") or "",
            ))
        for r in backup["blacklist"]:
            db.session.add(Blacklist(
                id=r["id"], name=r.get("name", "") or "",
                student_id=r.get("student_id", "") or "",
                reason=r.get("reason", "") or "",
                added_at=parse_dt(r.get("added_at")),
            ))
        db.session.commit()

        # 7. 统计
        print("\n=== 迁移完成 ===")
        print(f"  members: {Member.query.count()}")
        print(f"  projects: {Project.query.count()}")
        print(f"  activities: {Activity.query.count()}")
        print(f"  activity_participants: {ActivityParticipant.query.count()}")
        print(f"  irregular_activities: {IrregularActivity.query.count()}")
        print(f"  irregular_participants: {IrregularParticipant.query.count()}")
        print(f"  position_records: {PositionRecord.query.count()}")


if __name__ == "__main__":
    migrate()
