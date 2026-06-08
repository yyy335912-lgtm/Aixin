"""初始化数据库 + 爱心社示例数据"""
import os
import random
from datetime import datetime

from flask import Flask

from models import (
    db, Member, Project, Activity, ActivityParticipant,
    PositionRecord, IrregularActivity, IrregularParticipant,
    Blacklist, Setting, LongTermGroup,
)


PROJECTS = [
    (
        "心暖陇原",
        "书香润心，梦启陇原。2024年全新开设的品牌项目，与上海多阅公益合作，"
        "精准对接甘肃瓜州县腰站子东乡族镇中心小学、沙河乡中心小学。"
        "以远程阅读陪伴为核心，8 周一周期，志愿者与乡村孩子共读经典。",
        1,
    ),
    (
        "心暖苗疆",
        "暖心助学品牌。从「魔法冬衣」爱心捐赠到「倾听大山的声音」书信结对，"
        "陪伴贵州省台江县的孩子们成长，用冬衣抵御严寒、书信架起桥梁。",
        2,
    ),
    (
        "爱教三部曲",
        "社团深耕数十年的经典助学品牌：「大手牵小手→暑期支教→回访调研」闭环链路，"
        "深耕浙江武义、江西泰和、广西大化等地，让爱心从「结对」到「陪伴」再到「落地」。",
        3,
    ),
    (
        "心暖晨曦",
        "聚焦浙大后勤员工子女，一对一结对提供学业指导、课外素拓、情感互动等全方位陪伴，"
        "让求是园的爱心在校园内部传递。",
        4,
    ),
    (
        "心暖辰星",
        "以特殊儿童为关爱对象，针对其特点设计专属的陪伴活动与趣味课程，"
        "用温柔的陪伴为特殊孩童播种星星般的未来。",
        5,
    ),
    (
        "心暖桐心",
        "「凤凰鸣矣，于彼高岗；梧桐生矣，于彼朝阳」。"
        "以浙大退休老教师为陪伴对象，传承多年的经典暖心项目，实现跨代际的真挚交流。",
        6,
    ),
    (
        "心暖夕阳",
        "聚焦周边社区老人与敬老院老人，开展手工陪伴、节日慰问、谈心交流等多样活动，"
        "用年轻人的活力与真诚，为老人们的晚年生活增添色彩。",
        7,
    ),
    (
        "心暖杏林",
        "以幼儿健康为核心关爱方向，组织志愿者开展幼儿健康科普、爱心陪伴、趣味义诊等活动，"
        "用专业知识为孩子们的健康成长保驾护航。",
        8,
    ),
]

DEPARTMENTS = [
    "人资部", "综管部", "文体部", "活动部",
    "宣传部", "创设部", "外联部", "财务部",
]

LONG_TERM_GROUPS = [
    ("心暖陇原", "甘肃瓜州县腰站子东乡族镇中心小学、沙河乡中心小学——远程阅读陪伴", "#ff6b9d", 1),
    ("心暖苗疆", "贵州省台江县——魔法冬衣/倾听大山的声音 书信结对", "#9775fa", 2),
    ("心暖晨曦", "浙大后勤员工子女——一对一学业指导/课外素拓", "#ffa94d", 3),
    ("心暖辰星", "特殊儿童——专属陪伴活动与趣味课程", "#4dabf7", 4),
    ("心暖桐心", "浙大退休老教师——跨代际真挚交流", "#51cf66", 5),
    ("心暖夕阳", "周边社区老人/敬老院——手工陪伴/节日慰问", "#ffd43b", 6),
    ("心暖杏林", "幼儿健康——健康科普/爱心陪伴/趣味义诊", "#e64980", 7),
]

GRADES = ["2022级", "2023级", "2024级", "2025级", "2026级"]

SAMPLE_NAMES = [
    "陈郑熠", "符梦苑", "古丽吉乃提·艾沙", "郭兴宇", "李赛琳", "刘晨旭",
    "聂子尧", "彭乐嫣", "严静雯", "杨佳承", "杨智博", "李岩伟", "刘冀",
    "沈华", "安笑曼", "摆发英", "蔡璐璐", "曹莫铭", "陈晨", "陈侍羿",
    "陈显哲", "戴乔惠", "董嘉晟", "冯卓卡", "甘博", "高茂棋", "顾晨阳",
    "韩雨萱", "何雨欣", "胡可", "黄睿", "蒋雨桐", "李雨泽", "林思远",
    "刘思源", "陆创辉", "罗怡欣", "马晓宇", "钱思涵", "任英", "宋宇轩",
    "孙文博", "唐诗韵", "汪睿", "王梓涵", "吴佳怡", "夏语桐", "徐子轩",
    "许文博", "杨晨", "叶知秋", "余乐怡", "张驰", "张皓", "张雨欣",
    "赵思琦", "郑雨桐", "周雨欣", "朱锦仪", "朱一鸣", "邹雨泽",
]


def create_app():
    app = Flask(__name__)
    db_path = os.path.join(os.path.dirname(__file__), "data", "aixin.db")
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SECRET_KEY"] = "zju-aixin-secret-2026"
    db.init_app(app)
    return app


def init_db():
    app = create_app()
    with app.app_context():
        db.drop_all()
        db.create_all()

        # 1. 项目
        for name, desc, order in PROJECTS:
            db.session.add(Project(
                name=name, description=desc,
                sort_order=order, active=True,
            ))

        # 2. 系统设置
        db.session.add(Setting(key="admin_password", value="aixin2026"))
        db.session.add(Setting(key="current_semester", value="25秋冬"))

        # 3. 长期项目组
        long_term_groups = []
        for name, desc, color, order in LONG_TERM_GROUPS:
            g = LongTermGroup(
                name=name, description=desc, color=color,
                sort_order=order, active=True,
            )
            long_term_groups.append(g)
            db.session.add(g)
        db.session.commit()

        # 4. 成员（社员 + 长期项目组成员）
        members = []
        for i, name in enumerate(SAMPLE_NAMES):
            grade = GRADES[i % len(GRADES)]
            dept = DEPARTMENTS[i % len(DEPARTMENTS)]
            # 前 14 人设为长期项目组成员（按 7 个心暖系列组轮换），其余为社员
            if i < 14:
                identity = "长期项目组成员"
                group = long_term_groups[i % len(long_term_groups)]
            else:
                identity = "社员"
                # 部分社员也属于长期项目组（双重身份）
                if i % 3 == 0 and i < 30:
                    group = long_term_groups[i % len(long_term_groups)]
                else:
                    group = None
            m = Member(
                name=name, grade=grade, department=dept,
                student_id=f"223{str(i+1).zfill(8)}",
                identity=identity,
                long_term_group_id=group.id if group else None,
                phone=f"138{str(random.randint(10000000, 99999999))}",
            )
            members.append(m)
            db.session.add(m)
        db.session.commit()

        # 4. 活动（前 4 个项目各造 2 场，部分活动为 A+B 混合）
        projects = Project.query.order_by(Project.sort_order).all()
        activities = []
        for p in projects[:4]:
            for j in range(2):
                # 模拟：j=0 全是 A 类；j=1 A+B 都有
                if j == 0:
                    a_h, b_h = round(random.uniform(2.0, 4.0), 1), 0.0
                else:
                    a_h = round(random.uniform(1.0, 2.0), 1)
                    b_h = round(random.uniform(1.0, 2.0), 1)
                a = Activity(
                    project_id=p.id,
                    sequence=len(activities) + 1,
                    date=f"2025-1{(j % 9) + 1}-1{j+5}",
                    time=f"1{4 + j}:00-1{6 + j}:00",
                    location=f"{p.name}志愿服务点",
                    content=f"第{j+1}次常规活动",
                    served_count=10 + j * 5,
                    a_hours=a_h,
                    b_hours=b_h,
                    semester="25秋冬",
                    created_by="系统初始化",
                )
                activities.append(a)
                db.session.add(a)
        db.session.commit()

        # 5. 给活动添加参与者（每人随机 A / B 时长）
        for a in activities:
            n = random.randint(3, 6)
            for m in members[:n]:
                # 部分人只有 A 类，部分人有 A+B
                if random.random() < 0.6:
                    ap_a = round(random.uniform(1.0, 3.0), 1)
                    ap_b = 0.0
                else:
                    ap_a = round(random.uniform(0.5, 1.5), 1)
                    ap_b = round(random.uniform(0.5, 1.5), 1)
                db.session.add(ActivityParticipant(
                    activity_id=a.id, member_id=m.id,
                    a_hours=ap_a, b_hours=ap_b,
                ))
        db.session.commit()

        # 6. 任职积分
        positions = [
            ("社长", "", 20),
            ("副社长", "", 15),
            ("人资部部长", "人资部", 12),
            ("人资部副部长", "人资部", 8),
            ("综管部部长", "综管部", 12),
            ("活动部部长", "活动部", 12),
            ("宣传部部长", "宣传部", 12),
            ("创设部部长", "创设部", 12),
        ]
        for i, (pos, dept, pts) in enumerate(positions):
            m = members[i % len(members)]
            db.session.add(PositionRecord(
                member_id=m.id, position=pos, department=dept,
                points=pts, semester="25秋冬",
            ))

        # 7. 非常规活动（事件 + 参与者）
        irregulars_data = [
            ("参加唐奖颁奖典礼志愿服务", "2025-11-15", ["A"]),
            ("社团招新面试官", "2025-09-08", ["B"]),
            ("协助社团年检材料整理", "2025-12-05", ["A", "B"]),
            ("心暖苗疆魔法冬衣整理打包", "2025-12-20", ["A"]),
        ]
        for content, date, member_tags in irregulars_data:
            ia = IrregularActivity(
                content=content, date=date, semester="25秋冬",
            )
            db.session.add(ia)
            db.session.flush()
            # 选一些社员作为参与者
            for i, tag in enumerate(member_tags):
                m = members[(5 + i * 7) % len(members)]
                if tag == "A":
                    db.session.add(IrregularParticipant(
                        irregular_id=ia.id, member_id=m.id,
                        a_hours=2 + i, b_hours=0.0,
                    ))
                else:
                    db.session.add(IrregularParticipant(
                        irregular_id=ia.id, member_id=m.id,
                        a_hours=0.0, b_hours=2 + i,
                    ))

        # 8. 黑名单（占位）
        db.session.add(Blacklist(
            name="示例黑名单", student_id="0000000000",
            reason="示例数据，请删除",
        ))

        db.session.commit()
        print("=" * 50)
        print("✅ 数据库初始化完成！")
        print(f"   - 项目: {Project.query.count()} 个")
        print(f"   - 社员: {Member.query.count()} 人")
        print(f"   - 活动: {Activity.query.count()} 次")
        print(f"   - 活动参与记录: {ActivityParticipant.query.count()} 条")
        print(f"   - 非常规活动: {IrregularActivity.query.count()} 个")
        print(f"   - 非常规参与者: {IrregularParticipant.query.count()} 条")
        print(f"   - 任职记录: {PositionRecord.query.count()} 条")
        print("=" * 50)
        print("🔑 管理员密码: aixin2026 (可在「系统设置」中修改)")
        print("=" * 50)


if __name__ == "__main__":
    init_db()
