EMART STORE OPENING OS — v2.1 (蓝色专业版 / 9 页面)

WINDOWS 运行方式
1. 安装 Python 3.10+  https://www.python.org/downloads/
2. 解压 ZIP 到一个文件夹，在该文件夹打开 Command Prompt
3. py -m pip install -r requirements.txt
4. py -m streamlit run app.py
5. 浏览器打开 http://localhost:8501

9 个页面
Overview Dashboard / Opening Checklist / Department Progress / Project Timeline /
Site Visit Records / Issue & Risk Tracker / PIC & Workload / Photo Evidence / Data & Reports

说明
- 首次启动自动读取 Test_Dashboard(2).xlsx；之后的修改保存在 data/store_opening_master.xlsx
- 在 “Data & Reports” 页导入新的 Excel（须含名为 Checklist 的工作表）
- 主题颜色在 .streamlit/config.toml 和 app.py 顶部 (BLUE / NAVY …) 修改
- 照片保存在 data/photos，备份时请一并复制
- 多人同时使用请部署一个共享实例（内部服务器/云端），不要每人各跑一份
