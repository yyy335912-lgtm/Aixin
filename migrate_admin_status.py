"""一次性迁移：admin 表加 status / application_message / approved_at / approved_by
- 不破坏现有数据
- 所有现有 admin 自动设为 status='active'
"""
import os, sys
from sqlalchemy import inspect, text

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from app import app
from models import db, Admin


def migrate():
    with app.app_context():
        insp = inspect(db.engine)
        cols = [c["name"] for c in insp.get_columns("admins")]
        print("现有 admins 列:", cols)

        new_cols = {
            "status": "VARCHAR(16) DEFAULT 'active'",
            "application_message": "TEXT DEFAULT ''",
            "approved_at": "DATETIME",
            "approved_by": "VARCHAR(64) DEFAULT ''",
        }
        with db.engine.connect() as conn:
            for col, ddl in new_cols.items():
                if col not in cols:
                    print(f"  ADD COLUMN {col} {ddl}")
                    conn.execute(text(f"ALTER TABLE admins ADD COLUMN {col} {ddl}"))
                else:
                    print(f"  {col} 已存在，跳过")
            # 老数据全部设为 active（password_hash 非空即为活跃）
            conn.execute(text(
                "UPDATE admins SET status='active' "
                "WHERE status IS NULL OR status='' OR password_hash != ''"
            ))
            # 索引
            conn.execute(text("CREATE INDEX IF NOT EXISTS ix_admins_status ON admins(status)"))
            conn.commit()
        db.session.commit()
        print("=== admins 数据 ===")
        for a in Admin.query.all():
            print(f"  [{a.id}] {a.name} status={a.status} is_super={a.is_super_admin} is_root={a.is_root} has_pw={bool(a.password_hash)}")
        print("迁移完成")


if __name__ == "__main__":
    migrate()
