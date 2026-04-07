import os
from app import app

if __name__ == "__main__":
    # Render 会通过环境变量 PORT 告诉我们用哪个端口
    port = int(os.environ.get("PORT", 5000))
    # 在云端运行时，debug 必须设为 False，且 host 设为 '0.0.0.0'
    app.run(host='0.0.0.0', port=port, debug=False)