#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
新人数据自动化分析系统 - 启动脚本
双击此文件即可启动网页应用
"""

import subprocess
import sys
import os

# 切换到脚本所在目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("=" * 50)
print("  新人数据自动化分析系统")
print("=" * 50)
print()
print("正在启动网页应用...")
print("浏览器会自动打开 http://localhost:8501")
print()
print("按 Ctrl+C 可以停止服务")
print("=" * 50)
print()

# 启动Streamlit
subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", "8501"])

input("\n按回车键退出...")
