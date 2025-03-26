"""
运行Excel参数分析与可视化Web应用
"""

from app import app

if __name__ == '__main__':
    print('启动Excel参数分析与可视化Web应用...')
    print('在浏览器中访问: http://127.0.0.1:5000')
    app.run(debug=True)