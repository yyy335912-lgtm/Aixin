#!/bin/bash
# 一键启动浙江大学学生爱心社积分系统
# 使用方法：在终端执行  ./start.sh  或  bash start.sh

cd "$(dirname "$0")"

# 选择 Python 解释器（优先 anaconda3，其次 python3）
if [ -x "/opt/anaconda3/bin/python3" ]; then
    PYTHON="/opt/anaconda3/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
    PYTHON="python3"
else
    echo "❌ 未找到 Python 3，请先安装 Python 3.8+"
    exit 1
fi

echo "💗 使用 Python: $PYTHON"
echo "💗 版本: $($PYTHON --version)"

# 检查并安装依赖
echo ""
echo "📦 检查依赖..."
$PYTHON -c "import flask, flask_sqlalchemy, openpyxl" 2>/dev/null || {
    echo "📦 正在安装依赖..."
    $PYTHON -m pip install -r requirements.txt || {
        echo "❌ 依赖安装失败，请手动执行: pip3 install -r requirements.txt"
        exit 1
    }
}

# 初始化数据库（如果不存在）
if [ ! -f "data/aixin.db" ]; then
    echo ""
    echo "🗄️  首次运行，正在初始化数据库..."
    $PYTHON init_db.py
fi

# 启动 Flask
echo ""
echo "🚀 启动服务..."
$PYTHON app.py
